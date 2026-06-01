#!/usr/bin/env node
'use strict';

const { spawnSync } = require('child_process');
const { getBinaryPath } = require('../index.js');

const result = spawnSync(getBinaryPath(), process.argv.slice(2), {
  stdio: 'inherit',
});

if (result.error) {
  throw result.error;
}

if (result.signal) {
  process.kill(process.pid, result.signal);
}

process.exit(result.status ?? 1);
