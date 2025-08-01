window.addEventListener('DOMContentLoaded', () => {
  console.log('✅ Main page DOM loaded');

  // ===== Element References =====
  const folderPathInput = document.getElementById('folderPath');
  const browseBtn = document.getElementById('browseBtn');
  const runBtn = document.getElementById('runBtn');
  const progressText = document.getElementById('progressText');
  const afterExtraction = document.getElementById('afterExtraction');
  const checkResultsBtn = document.getElementById('checkResultsBtn');

  // ===== Folder Picker =====
  browseBtn.addEventListener('click', async () => {
    if (!window.electronAPI) {
      console.error('❌ electronAPI not available!');
      alert('electronAPI not loaded. Please restart the application.');
      return;
    }

    try {
      const folderPath = await window.electronAPI.selectFolder();
      if (folderPath) folderPathInput.value = folderPath;
    } catch (err) {
      console.error('❌ Error selecting folder:', err);
      alert('Error selecting folder: ' + err.message);
    }
  });

  // ===== Run Extraction =====
  runBtn.addEventListener('click', async () => {
    const folderPath = folderPathInput.value;
    if (!folderPath) return alert('Please select a folder first!');

    runBtn.disabled = true;
    progressText.textContent = '🔄 Extracting... Please Wait!';
    progressText.style.display = 'block';
    afterExtraction.style.display = 'none';

    try {
      const result = await window.electronAPI.runExtraction(folderPath);

      if (result.success) {
        progressText.textContent = '✅ Extraction Completed Successfully!';
        afterExtraction.style.display = 'block';
      } else {
        progressText.textContent = `❌ Error: ${result.error}`;
      }
    } catch (err) {
      progressText.textContent = `❌ Error: ${err.message}`;
    } finally {
      runBtn.disabled = false;
    }
  });

  // ===== Check Results =====
  checkResultsBtn.addEventListener('click', async () => {
    const folderPath = folderPathInput.value;
    if (!folderPath) return alert('Please select a folder first!');

    try {
      const result = await window.electronAPI.loadResults(folderPath);

      if (result.success) {
        window.electronAPI.showResultsPage(folderPath);
      } else {
        alert(`❌ Error loading results: ${result.error}`);
      }
    } catch (err) {
      alert(`❌ Error: ${err.message}`);
    }
  });
});