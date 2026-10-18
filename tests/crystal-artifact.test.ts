import test from 'node:test';
import assert from 'node:assert/strict';
import { inspectCrystalArtifact } from '../src/crystal-artifact.ts';

function fixture(version = 3) {
  const size = version === 3 ? 512 : 4096,
    b = Buffer.alloc(size * 3);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(b);
  b.writeUInt16LE(0x3e, 24);
  b.writeUInt16LE(version, 26);
  b.writeUInt16LE(0xfffe, 28);
  b.writeUInt16LE(version === 3 ? 9 : 12, 30);
  b.writeUInt16LE(6, 32);
  b.writeUInt32LE(version === 3 ? 0 : 1, 40);
  b.writeUInt32LE(1, 44);
  b.writeUInt32LE(1, 48);
  b.writeUInt32LE(4096, 56);
  b.writeUInt32LE(0xfffffffe, 60);
  b.writeUInt32LE(0xfffffffe, 68);
  b.fill(0xff, 76, 512);
  b.writeUInt32LE(0, 76);
  b.fill(0xff, size, 2 * size);
  b.writeUInt32LE(0xfffffffd, size);
  b.writeUInt32LE(0xfffffffe, size + 4);
  const entry = (id: number, name: string, type: number) => {
    const off = 2 * size + 128 * id;
    b.write(name + '\0', off, 'utf16le');
    b.writeUInt16LE((name.length + 1) * 2, off + 64);
    b[off + 66] = type;
    b[off + 67] = 1;
    for (const ptr of [68, 72, 76]) b.writeUInt32LE(0xffffffff, off + ptr);
    b.writeUInt32LE(0xfffffffe, off + 116);
    return off;
  };
  const root = entry(0, 'Root Entry', 5),
    first = entry(1, 'ReportInfo', 2);
  entry(2, 'QESession', 2);
  b.writeUInt32LE(1, root + 76);
  b.writeUInt32LE(2, first + 72);
  return b;
}

test('CFB v3/v4 inventory preserves directory evidence without claiming migration', () => {
  for (const version of [3, 4]) {
    const b = fixture(version),
      result = inspectCrystalArtifact(b);
    assert.equal(result.version, version);
    assert.equal(result.entries.length, 2);
    assert.deepEqual(
      result.entries.map((e) => e.path),
      [['ReportInfo'], ['QESession']],
    );
    assert.equal(result.crystalCandidate, true);
    assert.equal(result.automaticConversion, false);
    assert.match(result.sha256, /^[a-f0-9]{64}$/);
    assert.equal(result.bytes, b.length);
    b[0] = 0;
    assert.equal(result.entries[0]?.path[0], 'ReportInfo');
  }
});

test('generic OLE file is not authenticated as Crystal', () => {
  const b = fixture();
  b.fill(0, 1152, 1172);
  b.write('Other\0', 1152, 'utf16le');
  b.writeUInt16LE(12, 1216);
  assert.equal(inspectCrystalArtifact(b).crystalCandidate, false);
});

const corruptions: [string, (b: Buffer) => void][] = [
  [
    'signature',
    (b) => {
      b[0] = 0;
    },
  ],
  ['version', (b) => b.writeUInt16LE(5, 26)],
  ['byte order', (b) => b.writeUInt16LE(0, 28)],
  [
    'reserved header',
    (b) => {
      b[8] = 1;
    },
  ],
  ['extended DIFAT', (b) => b.writeUInt32LE(1, 72)],
  ['FAT count', (b) => b.writeUInt32LE(0, 44)],
  ['v3 directory count', (b) => b.writeUInt32LE(1, 40)],
  ['FAT marker', (b) => b.writeUInt32LE(0, 512)],
  ['FAT outside file', (b) => b.writeUInt32LE(100, 76)],
  ['unused DIFAT', (b) => b.writeUInt32LE(0, 80)],
  ['directory cycle', (b) => b.writeUInt32LE(1, 516)],
  ['directory overlap', (b) => b.writeUInt32LE(0, 48)],
  ['directory outside file', (b) => b.writeUInt32LE(99, 48)],
  [
    'missing root',
    (b) => {
      b[1090] = 0;
    },
  ],
  ['invalid name length', (b) => b.writeUInt16LE(65, 1216)],
  ['nonterminated name', (b) => b.writeUInt16LE(88, 1172)],
  [
    'invalid type',
    (b) => {
      b[1218] = 4;
    },
  ],
  [
    'second root',
    (b) => {
      b[1218] = 5;
    },
  ],
  ['declared oversize', (b) => b.writeUInt32LE(1000000, 1272)],
  ['tree cycle', (b) => b.writeUInt32LE(1, 1224)],
  ['invalid reference', (b) => b.writeUInt32LE(50, 1224)],
  ['stream children', (b) => b.writeUInt32LE(2, 1228)],
  ['unreachable entry', (b) => b.writeUInt32LE(0xffffffff, 1224)],
];
for (const [name, mutate] of corruptions)
  test(`rejects ${name}`, () => {
    const b = fixture();
    mutate(b);
    assert.throws(() => inspectCrystalArtifact(b), /invalid CFB artifact/);
  });
test('rejects truncation, wrong input and oversized artifact', () => {
  for (const n of [0, 8, 511, 513, 1535])
    assert.throws(() => inspectCrystalArtifact(fixture().subarray(0, n)));
  assert.throws(() => inspectCrystalArtifact('rpt' as unknown as Uint8Array));
  assert.throws(() => inspectCrystalArtifact(new Uint8Array(64 * 1024 * 1024 + 1)), /budget/);
});
