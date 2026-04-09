const electronInstaller = require('electron-winstaller');
const path = require('path');

async function buildInstaller() {
  try {
    console.log('Creating single file installer Setup.exe...');
    await electronInstaller.createWindowsInstaller({
      appDirectory: path.join(__dirname, 'dist', 'word-automator-win32-x64'),
      outputDirectory: path.join(__dirname, 'dist', 'installer'),
      authors: 'Antigravity',
      exe: 'word-automator.exe',
      description: 'Desktop Application for Word Automating'
    });
    console.log('It worked! Installer is at dist/installer/word-automator-1.0.0 Setup.exe');
  } catch (e) {
    console.log(`No dice: ${e.message}`);
  }
}

buildInstaller();
