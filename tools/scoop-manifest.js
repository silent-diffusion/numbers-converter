// Prints the Scoop manifest for a release: node tools/scoop-manifest.js <version> <sha256 of the zip>
// The release workflow writes it to bucket/numbers-converter.json, which makes this repository a
// Scoop bucket: scoop bucket add numbers-converter https://github.com/silent-diffusion/numbers-converter
const [version, hash] = process.argv.slice(2);
if (!/^\d+\.\d+\.\d+$/.test(version || '') || !/^[0-9a-f]{64}$/.test(hash || '')) {
  console.error('Usage: node tools/scoop-manifest.js <version> <sha256>');
  process.exit(1);
}

const repo = 'https://github.com/silent-diffusion/numbers-converter';
const zip = (v) => `${repo}/releases/download/v${v}/NumbersConverter-${v}-win-x64.zip`;
const manifest = {
  version,
  description: 'Convert Apple Numbers (.numbers) spreadsheets to Excel, offline.',
  homepage: repo,
  license: 'MIT',
  notes: 'Scoop keeps Numbers Converter up to date: run "scoop update numbers-converter".',
  architecture: { '64bit': { url: zip(version), hash } },
  shortcuts: [['Numbers Converter.exe', 'Numbers Converter']],
  checkver: { github: repo },
  autoupdate: { architecture: { '64bit': { url: zip('$version') } } }
};
process.stdout.write(`${JSON.stringify(manifest, null, 4)}\n`);
