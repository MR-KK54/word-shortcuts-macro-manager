const { spawn } = require('child_process');
const http = require('http');
const assert = require('assert');
const path = require('path');

console.log('Starting API Test Suite for Word Toolkit Server...');

const PORT = 3999;
const serverProcess = spawn(process.execPath, ['server.js'], {
  cwd: __dirname,
  env: { ...process.env, PORT: PORT.toString() },
  stdio: ['ignore', 'pipe', 'pipe']
});

let isReady = false;

serverProcess.stdout.on('data', data => {
  const str = data.toString();
  console.log('[Server]:', str.trim());
  if (str.includes('Word Toolkit Server running')) {
    isReady = true;
  }
});

serverProcess.stderr.on('data', data => {
  console.error('[Server Stderr]:', data.toString().trim());
});

serverProcess.on('error', err => console.error('[Server Process Error]:', err));
serverProcess.on('exit', (code, signal) => console.log(`[Server Process Exited]: code=${code}, signal=${signal}`));

function request(method, reqPath, body = null) {
  return new Promise((resolve, reject) => {
    const dataString = body ? JSON.stringify(body) : '';
    const req = http.request({
      hostname: '127.0.0.1',
      port: PORT,
      path: reqPath,
      method: method,
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(dataString)
      }
    }, res => {
      let raw = '';
      res.on('data', chunk => raw += chunk);
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, data: JSON.parse(raw) });
        } catch (e) {
          resolve({ status: res.statusCode, raw });
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(dataString);
    req.end();
  });
}

async function runTests() {
  // Wait up to 10 seconds for server ready signal
  for (let i = 0; i < 20; i++) {
    if (isReady) break;
    await new Promise(r => setTimeout(r, 500));
  }

  try {
    // 1. Health Check
    console.log('\nTest 1: Health Check');
    const res1 = await request('GET', '/api/health');
    assert.strictEqual(res1.status, 200, 'Health check status should be 200');
    assert.strictEqual(res1.data.status, 'ok', 'Health check status should be ok');
    console.log('  ✅ PASSED: Health check verified');

    // 2. Create Macro in User Group
    console.log('\nTest 2: Create Macro with User-Given Storing Group');
    const res2 = await request('POST', '/api/macros', {
      group: 'Legal Automation',
      name: 'AddSignatureBlock',
      type: 'bas',
      code: 'Sub AddSignature()\n  Selection.TypeText "Sincerely, John Doe"\nEnd Sub'
    });
    assert.strictEqual(res2.status, 200);
    assert.strictEqual(res2.data.success, true);
    assert.strictEqual(res2.data.action, 'created');
    assert.strictEqual(res2.data.macro.group, 'Legal Automation');
    console.log('  ✅ PASSED: Macro created under group "Legal Automation"');

    // 3. Replace Existing Macro with Same Name in Group
    console.log('\nTest 3: Replace Macro if Already Present');
    const res3 = await request('POST', '/api/macros', {
      group: 'Legal Automation',
      name: 'AddSignatureBlock',
      type: 'bas',
      code: 'Sub AddSignature()\n  Selection.TypeText "Sincerely, Jane Smith"\nEnd Sub'
    });
    assert.strictEqual(res3.status, 200);
    assert.strictEqual(res3.data.action, 'replaced');
    assert.strictEqual(res3.data.macro.code.includes('Jane Smith'), true);
    console.log('  ✅ PASSED: Existing macro successfully replaced');

    // 4. List Macros & Filter by Group Name
    console.log('\nTest 4: Filter Macros by Storing Group Name');
    const res4 = await request('GET', '/api/macros?group=Legal%20Automation');
    assert.strictEqual(res4.status, 200);
    assert.strictEqual(res4.data.macros.length, 1);
    assert.strictEqual(res4.data.macros[0].name, 'AddSignatureBlock');
    console.log('  ✅ PASSED: Group filtering verified');

    // 5. System Export & Import
    console.log('\nTest 5: Export & Import System Data');
    const res5 = await request('GET', '/api/export');
    assert.strictEqual(res5.status, 200);
    assert.ok(Array.isArray(res5.data.macros));

    const res6 = await request('POST', '/api/import', {
      mode: 'replace',
      data: res5.data
    });
    assert.strictEqual(res6.status, 200);
    assert.strictEqual(res6.data.success, true);
    console.log('  ✅ PASSED: System Export & Import package verified');

    console.log('\n====================================================');
    console.log('🎉 ALL API TESTS PASSED SUCCESSFULLY!');
    console.log('====================================================');

  } catch (err) {
    console.error('\n❌ TEST FAILED:', err);
    process.exitCode = 1;
  } finally {
    serverProcess.kill();
  }
}

runTests();
