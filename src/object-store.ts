import {
  S3Client,
  PutObjectCommand,
  GetObjectCommand,
  DeleteObjectCommand,
  HeadBucketCommand,
} from '@aws-sdk/client-s3';
import { createHash } from 'node:crypto';
import { ServiceError } from './service-error.ts';
export interface ObjectStore {
  put(key: string, bytes: Buffer): Promise<void>;
  get(key: string, expectedBytes: number, digest: string): Promise<Buffer>;
  delete(key: string): Promise<void>;
  ready(): Promise<void>;
  close(): void;
}
export const digest = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
export class S3Objects implements ObjectStore {
  readonly client: S3Client;
  readonly bucket: string;
  constructor(config: { bucket: string; region?: string; endpoint?: string; forcePathStyle?: boolean }) {
    this.bucket = config.bucket;
    this.client = new S3Client({
      region: config.region ?? 'us-east-1',
      endpoint: config.endpoint,
      forcePathStyle: config.forcePathStyle ?? !!config.endpoint,
      maxAttempts: 2,
    });
  }
  private key(key: string) {
    if (!/^[a-f0-9-]{36}$/.test(key)) throw new Error('Invalid internal object key');
    return `imprint/${key}`;
  }
  async ready() {
    await this.client.send(new HeadBucketCommand({ Bucket: this.bucket }), {
      abortSignal: AbortSignal.timeout(10000),
    });
  }
  async put(key: string, bytes: Buffer) {
    if (bytes.length > 32 * 1024 * 1024) throw new ServiceError(413, 'LIMIT_EXCEEDED');
    await this.client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(key),
        Body: bytes,
        ContentType: 'application/octet-stream',
        ChecksumSHA256: createHash('sha256').update(bytes).digest('base64'),
      }),
      { abortSignal: AbortSignal.timeout(30000) },
    );
  }
  async get(key: string, expectedBytes: number, expectedDigest: string) {
    if (expectedBytes < 0 || expectedBytes > 32 * 1024 * 1024)
      throw new ServiceError(500, 'OBJECT_INTEGRITY');
    const response = await this.client.send(
      new GetObjectCommand({ Bucket: this.bucket, Key: this.key(key) }),
      { abortSignal: AbortSignal.timeout(30000) },
    );
    if (!response.Body) throw new ServiceError(502, 'OBJECT_UNAVAILABLE');
    const chunks: Buffer[] = [];
    let length = 0;
    try {
      for await (const value of response.Body as AsyncIterable<Uint8Array>) {
        const bytes = Buffer.from(value);
        length += bytes.length;
        if (length > expectedBytes) throw new ServiceError(502, 'OBJECT_INTEGRITY');
        chunks.push(bytes);
      }
    } finally {
      if ('destroy' in response.Body) (response.Body as { destroy(): void }).destroy();
    }
    const bytes = Buffer.concat(chunks);
    if (length !== expectedBytes || digest(bytes) !== expectedDigest)
      throw new ServiceError(502, 'OBJECT_INTEGRITY');
    return bytes;
  }
  async delete(key: string) {
    await this.client.send(new DeleteObjectCommand({ Bucket: this.bucket, Key: this.key(key) }), {
      abortSignal: AbortSignal.timeout(10000),
    });
  }
  close() {
    this.client.destroy();
  }
}
