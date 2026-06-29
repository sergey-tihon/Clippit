'use strict';

const path = require('path');

const PLATFORM_PACKAGES = {
  'win32-x64':     { pkg: '@sergey-tihon/clippit-bin-win32-x64',     bin: 'clippit.exe' },
  'win32-arm64':   { pkg: '@sergey-tihon/clippit-bin-win32-arm64',   bin: 'clippit.exe' },
  'darwin-x64':    { pkg: '@sergey-tihon/clippit-bin-darwin-x64',    bin: 'clippit'     },
  'darwin-arm64':  { pkg: '@sergey-tihon/clippit-bin-darwin-arm64',  bin: 'clippit'     },
  'linux-x64':     { pkg: '@sergey-tihon/clippit-bin-linux-x64',     bin: 'clippit'     },
  'linux-arm64':   { pkg: '@sergey-tihon/clippit-bin-linux-arm64',   bin: 'clippit'     },
};

/**
 * Returns the absolute path to the native clippit binary for the current platform.
 * Throws if the platform is unsupported or the platform package is not installed.
 */
function getBinaryPath() {
  const platformKey = `${process.platform}-${process.arch}`;
  const entry = PLATFORM_PACKAGES[platformKey];

  if (!entry) {
    throw new Error(`Unsupported platform: ${platformKey}`);
  }

  return require.resolve(`${entry.pkg}/${entry.bin}`);
}

module.exports = { getBinaryPath };
