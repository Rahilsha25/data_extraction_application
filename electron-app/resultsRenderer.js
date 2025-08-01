window.addEventListener('DOMContentLoaded', () => {
  console.log('✅ Results page loaded');

  const container = document.getElementById('resultsTableContainer');
  const homeBtn = document.getElementById('homeBtn');

  const waitForElectronAPI = () => {
    if (window.electronAPI) {
      setupEventListeners();
    } else {
      setTimeout(waitForElectronAPI, 100); // Retry every 100ms until available
    }
  };

  const setupEventListeners = () => {
    // Home button navigation
    if (homeBtn) {
      homeBtn.addEventListener('click', () => {
        if (window.electronAPI?.goHome) {
          window.electronAPI.goHome();
        } else {
          alert('Navigation error: goHome function not available. Please restart the application.');
        }
      });
    }

    // Handle folder path and show results
    window.electronAPI.onSetFolderPath(async (event, folderPath) => {
      if (!folderPath) {
        container.innerHTML = '<p class="error">❌ Folder path not found.</p>';
        return;
      }

      const result = await window.electronAPI.loadResults(folderPath);
      if (!result.success) {
        container.innerHTML = `<p class="error">❌ ${result.error}</p>`;
        return;
      }

      const data = result.data;
      if (!data || data.length === 0) {
        container.innerHTML = `<p class="error">⚠️ No extracted results found in Excel.</p>`;
        return;
      }

      renderResultsTable(data);
    });
  };

  const renderResultsTable = (data) => {
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Create header
    const headerRow = document.createElement('tr');
    ['Title', 'Answer', 'Reference'].forEach(header => {
      const th = document.createElement('th');
      th.innerText = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Add rows
    data.forEach(row => {
      const tr = document.createElement('tr');
      ['Title', 'Answer', 'Reference'].forEach(field => {
        const td = document.createElement('td');
        td.innerText = row[field] || '';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    container.innerHTML = '';
    container.appendChild(table);
  };

  waitForElectronAPI();
});