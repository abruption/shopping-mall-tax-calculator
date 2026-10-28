const { validateAudit } = require('../../scripts/validate-audit');

const report = () => ({
  advisories: { example: {
    github_advisory_id: 'GHSA-gcqx-w48m-w2h4', module_name: 'deepmerge', severity: 'low', patched_versions: null,
    findings: [{ version: '4.3.1', dev: true, bundled: false, paths: ['.>jest>@jest/core>jest-config>deepmerge'] }],
  } },
  metadata: { vulnerabilities: { info: 0, low: 1, moderate: 0, high: 0, critical: 0 } },
});

describe('visible, exact audit risk policy', () => {
  it('accepts clean reports in both modes', () => {
    const clean = { advisories: {}, metadata: { vulnerabilities: { info: 0, low: 0, moderate: 0, high: 0, critical: 0 } } };
    expect(validateAudit(clean, 'full').remaining).toBe(0);
    expect(validateAudit(clean, 'production').remaining).toBe(0);
  });
  it('reports the exact unresolved Jest advisory rather than calling the audit clean', () => {
    expect(validateAudit(report(), 'full')).toEqual({ remaining: 1, advisory: 'GHSA-gcqx-w48m-w2h4' });
  });
  it('rejects any production finding', () => expect(() => validateAudit(report(), 'production')).toThrow());
  it.each(['identity', 'version', 'path', 'production', 'severity', 'published fix', 'new low'])('rejects changed risk: %s', change => {
    const data: any = report(); const advisory = data.advisories.example;
    if (change === 'identity') advisory.github_advisory_id = 'GHSA-other';
    if (change === 'version') advisory.findings[0].version = '4.3.2';
    if (change === 'path') advisory.findings[0].paths = ['.>application>deepmerge'];
    if (change === 'production') advisory.findings[0].dev = false;
    if (change === 'severity') advisory.severity = 'high';
    if (change === 'published fix') advisory.patched_versions = '>=4.3.2';
    if (change === 'new low') { data.advisories.extra = { ...advisory, github_advisory_id: 'GHSA-new' }; data.metadata.vulnerabilities.low++; }
    expect(() => validateAudit(data, 'full')).toThrow();
  });
  it.each([{}, { advisories: 1, metadata: { vulnerabilities: { info: 0, low: 0, moderate: 0, high: 0, critical: 0 } } }, { error: 'registry unavailable' }, { advisories: {}, metadata: { vulnerabilities: { low: 1 } } }])('fails closed on incomplete reports', data => {
    expect(() => validateAudit(data, 'full')).toThrow();
  });
});
