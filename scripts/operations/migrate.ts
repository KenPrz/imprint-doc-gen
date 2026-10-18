import { configuredStore } from '../../src/config-store.ts';
import { migrateSqlite } from '../../src/migrate.ts';
const snapshot = process.argv[2];
if (!snapshot || process.argv.length !== 3)
  throw new Error('Usage: npm run migrate -- /path/to/offline-backup.sqlite');
const target = await configuredStore();
if (!target) throw new Error('DATABASE_URL and S3_BUCKET are required');
try {
  console.log(JSON.stringify(await migrateSqlite(snapshot, target), null, 2));
} finally {
  await target.close();
}
