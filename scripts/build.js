const { spawn } = require('child_process');
const fs = require('fs-extra');
const path = require('path');

async function build() {
  console.log('🧹 Cleaning dist directory...');
  await fs.emptyDir('dist');

  console.log('📦 Building with SWC...');
  
  // Build all TypeScript files with proper output structure
  await runCommand('swc', ['src', '-d', 'dist', '--config-file', '.swcrc', '--ignore', '**/__tests__/**,**/*.test.ts,**/*.spec.ts']);

  // Copy static files
  console.log('📄 Copying static files...');
  await fs.copy('src/renderer/index.html', 'dist/src/renderer/index.html');
  await fs.copy('src/locales', 'dist/locales');

  console.log('✅ Build completed!');
}

function runCommand(command, args) {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { 
      stdio: 'inherit',
      shell: process.platform === 'win32'
    });
    
    proc.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`${command} exited with code ${code}`));
      } else {
        resolve();
      }
    });
  });
}

build().catch(console.error);