import { createHash } from 'node:crypto';

/** Passive CFB v3/v4 directory inventory only. Stream names are observations,
 * not decoded Crystal formulas, layout, data or evidence of executable compatibility.
 * Extended DIFAT files are deliberately rejected until independently exercised.
 */
export function inspectCrystalArtifact(input: Uint8Array) {
  if (!(input instanceof Uint8Array) || input.byteLength > 64 * 1024 * 1024)
    throw new Error('Artifact must be bytes within the 64 MiB inventory budget');
  const data = Buffer.from(input);
  const fail: (reason: string) => never = (reason) => {
    throw new Error(`Unsupported or invalid CFB artifact: ${reason}`);
  };
  if (data.length < 512 || data.subarray(0, 8).toString('hex') !== 'd0cf11e0a1b11ae1')
    fail('signature/header');
  const version = data.readUInt16LE(26),
    shift = data.readUInt16LE(30);
  if (
    !((version === 3 && shift === 9) || (version === 4 && shift === 12)) ||
    data.readUInt16LE(28) !== 0xfffe ||
    data.readUInt16LE(32) !== 6
  )
    fail('version/byte order/sector size');
  if (data.subarray(8, 24).some(Boolean) || data.subarray(34, 40).some(Boolean))
    fail('reserved header bytes');
  const size = 2 ** shift,
    count = data.length / size - 1;
  if (!Number.isInteger(count) || count < 2) fail('truncated sectors');
  const fatCount = data.readUInt32LE(44),
    dirCount = data.readUInt32LE(40);
  if (version === 3 && dirCount !== 0) fail('v3 directory count');
  if (data.readUInt32LE(72) !== 0 || data.readUInt32LE(68) !== 0xfffffffe || fatCount > 109)
    fail('extended DIFAT is outside the supported inventory profile');
  if (fatCount < 1 || fatCount > count || data.readUInt32LE(56) !== 4096)
    fail('FAT count/mini-stream cutoff');
  const sector = (id: number) => {
    if (!Number.isInteger(id) || id < 0 || id >= count) fail('sector outside file');
    return data.subarray((id + 1) * size, (id + 2) * size);
  };
  const fatIds: number[] = [],
    fat: number[] = [];
  for (let i = 0; i < 109; i++) {
    const id = data.readUInt32LE(76 + i * 4);
    if (i >= fatCount) {
      if (id !== 0xffffffff) fail('unused DIFAT entry');
      continue;
    }
    if (fatIds.includes(id)) fail('duplicate FAT sector');
    fatIds.push(id);
    const bytes = sector(id);
    for (let j = 0; j < size; j += 4) fat.push(bytes.readUInt32LE(j));
  }
  if (fat.length < count || fatIds.some((id) => fat[id] !== 0xfffffffd)) fail('FAT coverage/marker');
  const visited = new Set<number>(),
    directory: Buffer[] = [];
  let id = data.readUInt32LE(48);
  while (id !== 0xfffffffe) {
    if (visited.has(id) || fatIds.includes(id)) fail('cyclic/overlapping directory chain');
    if (visited.size >= 1024) fail('directory sector budget');
    visited.add(id);
    directory.push(sector(id));
    id = fat[id] ?? fail('missing FAT entry');
  }
  if (!directory.length || (version === 4 && directory.length !== dirCount)) fail('directory count');
  const raw = Buffer.concat(directory);
  const entries: {
    id: number;
    name: string;
    kind: 'storage' | 'stream' | 'root';
    declaredBytes: number;
    left: number;
    right: number;
    child: number;
  }[] = [];
  for (let i = 0; i < raw.length; i += 128) {
    const entry = raw.subarray(i, i + 128),
      type = entry[66];
    if (type === 0) continue;
    if (![1, 2, 5].includes(type!) || (type === 5 && i !== 0)) fail('directory object type');
    const length = entry.readUInt16LE(64);
    if (length < 2 || length > 64 || length % 2 || entry.readUInt16LE(length - 2) !== 0)
      fail('directory name');
    const name = entry.subarray(0, length - 2).toString('utf16le');
    if (!name || name.includes('\0')) fail('directory name');
    const declared = version === 3 ? BigInt(entry.readUInt32LE(120)) : entry.readBigUInt64LE(120);
    if (declared > BigInt(data.length)) fail('declared stream size exceeds file');
    entries.push({
      id: i / 128,
      name,
      kind: type === 1 ? 'storage' : type === 2 ? 'stream' : 'root',
      declaredBytes: Number(declared),
      left: entry.readUInt32LE(68),
      right: entry.readUInt32LE(72),
      child: entry.readUInt32LE(76),
    });
  }
  const root = entries[0];
  if (
    !root ||
    root.id !== 0 ||
    root.kind !== 'root' ||
    root.name !== 'Root Entry' ||
    root.left !== 0xffffffff ||
    root.right !== 0xffffffff
  )
    fail('root directory');
  const byId = new Map(entries.map((e) => [e.id, e])),
    seen = new Set<number>([0]);
  const output: { directoryId: number; path: string[]; kind: 'storage' | 'stream'; declaredBytes: number }[] =
    [];
  const stack = [{ id: root.child, parent: [] as string[] }];
  while (stack.length) {
    const next = stack.pop()!;
    if (next.id === 0xffffffff) continue;
    if (seen.has(next.id)) fail('cyclic/shared directory reference');
    const entry = byId.get(next.id);
    if (!entry || entry.kind === 'root') fail('invalid directory reference');
    seen.add(entry.id);
    const path = [...next.parent, entry.name];
    if (path.length > 64) fail('storage nesting budget');
    output.push({ directoryId: entry.id, path, kind: entry.kind, declaredBytes: entry.declaredBytes });
    stack.push({ id: entry.left, parent: next.parent }, { id: entry.right, parent: next.parent });
    if (entry.kind === 'stream' && entry.child !== 0xffffffff) fail('stream has children');
    if (entry.kind === 'storage') stack.push({ id: entry.child, parent: path });
  }
  if (seen.size !== entries.length) fail('unreachable directory entry');
  output.sort((a, b) => a.directoryId - b.directoryId);
  return {
    schema: 'imprint-crystal-artifact-v1' as const,
    sha256: createHash('sha256').update(data).digest('hex'),
    bytes: data.length,
    container: 'cfb' as const,
    version,
    entries: output,
    crystalCandidate:
      output.some((e) => e.kind === 'stream' && e.path.at(-1) === 'ReportInfo') &&
      output.some((e) => e.kind === 'stream' && e.path.at(-1) === 'QESession'),
    automaticConversion: false as const,
    limitations: [
      'Directory metadata only; stream payloads and allocation chains are not validated or decoded.',
      'Candidate detection uses observed stream names, not an authenticated Crystal format signature.',
      'No formulas, layout, database credentials, records, or subreport semantics are extracted.',
      'Extended DIFAT containers are unsupported.',
    ],
  };
}
