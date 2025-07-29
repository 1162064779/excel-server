// volume-guard.js ------------------------------------------------------------
const { execSync } = require('child_process');
const path = require('path');

// ★★★ 关键改动：用 process.execPath 而不是 __dirname ★★★
function getExecutableRoot() {
  // 示例： F:\excel-server.exe → F:\
  return path.parse(process.execPath).root;
}

function getCurrentVolumeId() {
  const root = getExecutableRoot();

  try {
    if (process.platform === 'win32') {
      // wmic 需要类似 'F:' 这样的参数
      const drive = root.replace(/\\$/, '');
      const cmd   = `wmic volume where DriveLetter='${drive}' get SerialNumber /value`;
      const out   = execSync(cmd, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
      const m     = out.match(/SerialNumber=(\w+)/);
      return m ? m[1].trim() : null;
    }

    if (process.platform === 'darwin') {
      const cmd = `diskutil info "${root}"`;
      const out = execSync(cmd, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
      const m   = out.match(/Volume UUID:\s+([0-9A-F-]+)/i);
      return m ? m[1].trim().toUpperCase() : null;
    }
  } catch (_) {
    /* ignore */
  }

  return null;
}

/* ------------ 你的白名单 ------------ */
const ALLOWED = {
  win32 : new Set(['1858550701', '6EE4-71BD']),          // F: 盘
  darwin: new Set(['6AA1816A-3D68-39CE-A0C9-12408C26D2F9']) // macOS
};
/* ------------------------------------ */

(function guard () {
  const id    = getCurrentVolumeId();
  const white = ALLOWED[process.platform] || new Set();

  if (!id || !white.has(id)) {
    console.error(`
====================================================
   本程序仅限在授权 U 盘启动！
   当前卷 ID：${id || '未知'}
   允许 ID  ：${[...white].join(', ') || '无'}
   程序即将退出...
====================================================`);
    process.exit(1);
  }
})();