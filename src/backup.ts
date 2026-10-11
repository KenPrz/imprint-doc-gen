import Database from 'better-sqlite3';
import { access, chmod, copyFile, mkdir, readdir, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { constants } from 'node:fs';
import { randomUUID } from 'node:crypto';

function verify(path: string) {
  const db = new Database(path, { readonly: true, fileMustExist: true });
  try {
    if (db.pragma('integrity_check', { simple: true }) !== 'ok')
      throw new Error('Backup integrity check failed');
    if (db.pragma('user_version', { simple: true }) !== 1) throw new Error('Unsupported backup schema');
    if ((db.pragma('foreign_key_check') as unknown[]).length)
      throw new Error('Backup has invalid foreign keys');
    for (const table of ['tenants', 'templates', 'versions', 'jobs'])
      db.prepare(`SELECT COUNT(*) FROM ${table}`).get();
  } finally {
    db.close();
  }
}
export async function backupSqlite(directory: string, destination: string) {
  const source = resolve(directory, 'imprint.sqlite'),
    target = resolve(destination);
  if (source === target) throw new Error('Backup must not overwrite the active database');
  try {
    await access(target);
    throw new Error('Backup destination already exists');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  }
  await mkdir(dirname(target), { recursive: true, mode: 0o700 });
  const temporary = `${target}.${randomUUID()}.tmp`,
    db = new Database(source, { readonly: true, fileMustExist: true });
  try {
    await db.backup(temporary);
    await chmod(temporary, 0o600);
    verify(temporary);
    // COPYFILE_EXCL also handles a competing backup process creating the destination.
    await copyFile(temporary, target, constants.COPYFILE_EXCL);
  } finally {
    db.close();
    await rm(temporary, { force: true });
  }
}
export async function restoreSqlite(backup: string, directory: string) {
  verify(backup);
  await mkdir(directory, { recursive: true, mode: 0o700 });
  if ((await readdir(directory)).length)
    throw new Error('Restore requires an empty directory and a stopped service');
  const target = join(directory, 'imprint.sqlite');
  await copyFile(backup, target, constants.COPYFILE_EXCL);
  await chmod(target, 0o600);
  verify(target);
}
