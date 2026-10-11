import { PostgresStore } from './postgres-store.ts';
import { S3Objects } from './object-store.ts';
export async function configuredStore() {
  if (!process.env.DATABASE_URL) return;
  if (!process.env.S3_BUCKET) throw new Error('S3_BUCKET is required with DATABASE_URL');
  const store = new PostgresStore(
    process.env.DATABASE_URL,
    new S3Objects({
      bucket: process.env.S3_BUCKET,
      region: process.env.AWS_REGION,
      endpoint: process.env.S3_ENDPOINT,
    }),
  );
  try {
    await store.initialize();
    return store;
  } catch (error) {
    await store.close();
    throw error;
  }
}
