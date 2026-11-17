const vscode = require('vscode');
const { randomBytes } = require('node:crypto');
const { viewerHtml } = require('./html');

function activate(context) {
  context.subscriptions.push(vscode.window.registerCustomEditorProvider('lumen-pptx.preview', {
    async openCustomDocument(uri) { return { uri, dispose() {} }; },
    async resolveCustomEditor(document, panel) {
      const media = vscode.Uri.joinPath(context.extensionUri, 'media');
      panel.webview.options = { enableScripts: true, localResourceRoots: [media] };
      let disposed = false, generation = 0, timer;
      const send = async () => {
        const request = ++generation;
        try {
          const stat = await vscode.workspace.fs.stat(document.uri);
          if (stat.size > 100 * 1024 * 1024) throw new Error('This presentation exceeds the 100 MB preview limit.');
          const bytes = await vscode.workspace.fs.readFile(document.uri);
          if (!disposed && request === generation) {
            await panel.webview.postMessage({ type: 'document', data: Buffer.from(bytes).toString('base64') });
          }
        } catch (error) {
          if (!disposed) await panel.webview.postMessage({ type: 'error', text: error.message });
        }
      };
      const messages = panel.webview.onDidReceiveMessage(message => {
        if (message.type === 'ready' || message.type === 'reload') void send();
        if (message.type === 'link' && typeof message.url === 'string') {
          try {
            const uri = vscode.Uri.parse(message.url);
            if (['https', 'http', 'mailto'].includes(uri.scheme)) void vscode.env.openExternal(uri);
          } catch {}
        }
      });
      // Escape glob syntax so names like "deck [v2].pptx" match literally.
      const fileName = document.uri.path.split('/').pop().replace(/[[\]{}*?!]/g, '[$&]');
      const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
        vscode.Uri.joinPath(document.uri, '..'), fileName
      ));
      const reload = () => { clearTimeout(timer); timer = setTimeout(send, 250); };
      const changed = watcher.onDidChange(reload);
      // Apps that save by replacing the file fire delete then create; reload when it reappears.
      const created = watcher.onDidCreate(reload);
      const deleted = watcher.onDidDelete(() => panel.webview.postMessage({ type: 'error', text: 'The source PPTX was deleted or moved.' }));
      panel.onDidDispose(() => {
        disposed = true; generation++; clearTimeout(timer);
        messages.dispose(); changed.dispose(); created.dispose(); deleted.dispose(); watcher.dispose();
      });
      panel.webview.html = viewerHtml({
        cspSource: panel.webview.cspSource,
        nonce: randomBytes(24).toString('hex'),
        asset: name => panel.webview.asWebviewUri(vscode.Uri.joinPath(media, name)).toString()
      });
    }
  }, { webviewOptions: { retainContextWhenHidden: true }, supportsMultipleEditorsPerDocument: true }));
}
module.exports = { activate };
