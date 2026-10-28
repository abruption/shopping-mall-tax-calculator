const fs = require('fs');

const knownAdvisory = 'GHSA-gcqx-w48m-w2h4';
const knownPaths = new Set([
  '.>jest>@jest/core>jest-config>deepmerge',
  '.>jest>jest-cli>@jest/core>jest-config>deepmerge',
  '.>jest>jest-cli>jest-config>deepmerge',
  '.>ts-jest>jest>@jest/core>jest-config>deepmerge',
  '.>ts-jest>jest>jest-cli>@jest/core>jest-config>deepmerge',
  '.>ts-jest>jest>jest-cli>jest-config>deepmerge',
]);

function validateAudit(report, mode) {
  if (!['full', 'production'].includes(mode)) throw new Error('Unknown audit mode');
  if (!report || report.error || !report.advisories || typeof report.advisories !== 'object' || Array.isArray(report.advisories)) {
    throw new Error('Audit failed or returned an unsupported report');
  }
  const counts = report.metadata?.vulnerabilities;
  const severities = ['info', 'low', 'moderate', 'high', 'critical'];
  if (!counts || severities.some(key => !Number.isInteger(counts[key]) || counts[key] < 0)) {
    throw new Error('Audit vulnerability counts are missing or invalid');
  }
  const advisories = Object.values(report.advisories);
  if (severities.reduce((total, key) => total + counts[key], 0) !== advisories.length) {
    throw new Error('Audit counts do not match its advisory list');
  }
  if (advisories.length === 0) return { remaining: 0 };
  if (mode === 'production' || advisories.length !== 1) {
    throw new Error('Unexpected dependency advisory: review and fix before accepting');
  }
  const advisory = advisories[0];
  if (advisory.github_advisory_id !== knownAdvisory || advisory.module_name !== 'deepmerge' ||
      advisory.severity !== 'low' || counts.low !== 1 || advisory.patched_versions !== null ||
      !Array.isArray(advisory.findings) || advisory.findings.length === 0 ||
      advisory.findings.some(finding => finding.version !== '4.3.1' || finding.dev !== true ||
        finding.bundled !== false || !Array.isArray(finding.paths) || finding.paths.length === 0 ||
        finding.paths.some(path => !knownPaths.has(path)))) {
    throw new Error('Advisory differs from the documented unpatched development-only Jest risk');
  }
  return { remaining: 1, advisory: knownAdvisory };
}

if (require.main === module) {
  try {
    const result = validateAudit(JSON.parse(fs.readFileSync(process.argv[2], 'utf8')), process.argv[3]);
    console.log(result.remaining === 0 ? 'Audit: no findings.' :
      `Audit: ${result.advisory} remains LOW, unpatched, development-only through Jest. NOT FIXED; full report retained.`);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
module.exports = { validateAudit };
