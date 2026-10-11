import { backupSqlite, restoreSqlite } from '../../src/backup.ts';
const [action, source, target] = process.argv.slice(2);
if (!source || !target || !['backup', 'restore'].includes(action ?? ''))
  throw new Error(
    'Usage: node scripts/operations/backup.ts backup <data-dir> <backup.sqlite> | restore <backup.sqlite> <empty-data-dir>',
  );
await (action === 'backup' ? backupSqlite(source, target) : restoreSqlite(source, target));
console.log(`${action} verified: ${target}`);
