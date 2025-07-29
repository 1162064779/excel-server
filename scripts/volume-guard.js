const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');

/* ---------- ① 找到磁盘上的真实可执行文件 -------------------------- */
function realExecPath () {
  const ep = process.pkg?.entrypoint;
  if (ep && !ep.startsWith('/snapshot/')) return fs.realpathSync(ep);
  return fs.realpathSync(process.execPath);
}

/* ---------- ② 针对平台取得卷 ID ---------------------------------- */
function getCurrentVolumeId () {
  if (process.platform === 'win32') {
    const drive = path.parse(realExecPath()).root.replace(/\\$/, ''); // "F:"
    try {
      const out = execSync(
        `wmic volume where DriveLetter='${drive}' get SerialNumber /value`,
        { encoding:'utf8', stdio:['ignore','pipe','ignore'] }
      );
      const m = out.match(/SerialNumber=(\w+)/);
      return m ? m[1].trim() : null;
    } catch { return null; }
  }

  if (process.platform === 'darwin') {
    try {
      /* 1) df -P -> 设备节点 */
      const df = execSync(
        `/bin/df -P "${realExecPath()}" | tail -1 | awk '{print $1}'`,
        { encoding:'utf8', stdio:['ignore','pipe','ignore'] }
      ).trim();                                   // 例：/dev/disk2s1

      /* 2) diskutil info 设备节点 → 解析 Volume UUID */
      const info = execSync(
        `/usr/sbin/diskutil info "${df}"`,
        { encoding:'utf8', stdio:['ignore','pipe','ignore'] }
      );
      const m = info.match(/Volume UUID:\s+([0-9A-F-]+)/i);
      return m ? m[1].toUpperCase() : null;
    } catch { return null; }
  }

  return null;   // 其它平台暂不支持
}

/* ---------- ③ 白名单 -------------------------------------------- */
const ALLOWED = {
  win32 : new Set(['1858550701', '6EE4-71BD']),
  darwin: new Set(['6AA1816A-3D68-39CE-A0C9-12408C26D2F9'])
};

/* ---------- ④ 守卫 ---------------------------------------------- */
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