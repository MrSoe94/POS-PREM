const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function getCertPaths(baseDir) {
  const certDir = path.join(baseDir, 'ssl');
  return {
    certDir,
    keyPath: path.join(certDir, 'key.pem'),
    certPath: path.join(certDir, 'cert.pem')
  };
}

function readExistingCert(keyPath, certPath) {
  try {
    if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) return null;
    return {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    };
  } catch {
    return null;
  }
}

async function generateWithSelfSigned(certDir, keyPath, certPath) {
  const selfsigned = require('selfsigned');
  const pems = await selfsigned.generate([{ name: 'commonName', value: 'POS-Local' }], {
    keySize: 2048,
    days: 3650,
    algorithm: 'sha256'
  });
  fs.mkdirSync(certDir, { recursive: true });
  fs.writeFileSync(keyPath, pems.private);
  fs.writeFileSync(certPath, pems.cert);
  return readExistingCert(keyPath, certPath);
}

async function ensureLocalHttpsCredentials(baseDir) {
  const { certDir, keyPath, certPath } = getCertPaths(baseDir);
  const existing = readExistingCert(keyPath, certPath);
  if (existing) return existing;

  try {
    fs.mkdirSync(certDir, { recursive: true });
    const subj = '/CN=POS-Local/O=POS/C=ID';
    execSync(
      `openssl req -x509 -newkey rsa:2048 -keyout "${keyPath}" -out "${certPath}" -days 3650 -nodes -subj "${subj}"`,
      { stdio: 'ignore' }
    );
    const fromOpenSsl = readExistingCert(keyPath, certPath);
    if (fromOpenSsl) return fromOpenSsl;
  } catch {}

  try {
    return await generateWithSelfSigned(certDir, keyPath, certPath);
  } catch {
    return null;
  }
}

module.exports = {
  ensureLocalHttpsCredentials,
  getCertPaths
};
