import * as fs from 'fs';
import * as path from 'path';

const root = path.resolve(__dirname, '../..');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const lock = fs.readFileSync(path.join(root, 'pnpm-lock.yaml'), 'utf8');

// Read package keys, not snapshots or override selectors. All entries are checked,
// including duplicate major lines used by different build/test tools.
const versions = (name: string): string[] => {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return [...lock.matchAll(new RegExp(`^  '?${escaped}@(\\d+\\.\\d+\\.\\d+)'?:$`, 'gm'))].map(match => match[1]);
};
const atLeast = (actual: string, floor: string): boolean => {
  const a = actual.split('.').map(Number), b = floor.split('.').map(Number);
  for (let i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] > b[i];
  return true;
};

describe('build dependency security baseline', () => {
  it('does not restore unused webpack loader or runtime dependency chains', () => {
    expect(manifest.devDependencies['swc-loader']).toBeUndefined();
    expect(versions('webpack')).toEqual([]);
    expect(versions('serialize-javascript')).toEqual([]);
    expect(Object.keys(manifest.dependencies)).toEqual(['xlsx']);
  });

  it.each([
    ['tmp', '0.2', '0.2.7'], ['glob', '10', '10.5.0'],
    ['js-yaml', '3', '3.15.2'], ['js-yaml', '4', '4.3.2'],
    ['minimatch', '3', '3.1.4'], ['minimatch', '5', '5.1.8'],
    ['picomatch', '2', '2.3.2'], ['lodash', '4', '4.18.1'],
    ['@xmldom/xmldom', '0.8', '0.8.15'], ['fast-uri', '3', '3.1.9'],
    ['browserslist', '4', '4.28.7'],
    ['brace-expansion', '1', '1.1.21'], ['brace-expansion', '2', '2.1.7'],
  ])('%s %s.x never resolves below %s', (name, line, floor) => {
    const resolved = versions(name).filter(version => version.startsWith(`${line}.`));
    expect(resolved.length).toBeGreaterThan(0);
    resolved.forEach(version => expect(atLeast(version, floor)).toBe(true));
  });
});
