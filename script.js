let sheetData = [];

        async function fetchSheetData() {
            const sheetId = document.getElementById('sheetId').value;
            const apiKey = document.getElementById('apiKey').value;
            
            if (!sheetId || !apiKey) {
                alert('Please enter both API Key and Google Sheet ID');
                return;
            }
            
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/A1:Z1000?key=${apiKey}`;
            
            try {
                const response = await fetch(url);
                const data = await response.json();
                
                if (data.values) {
                    sheetData = data.values;
                    alert('Data fetched successfully!');
                } else {
                    alert('Failed to fetch data. Check the Sheet ID, API Key, and permissions.');
                }
            } catch (error) {
                alert('Error fetching data: ' + error.message);
            }
        }

        function downloadExcel() {
            if (sheetData.length === 0) {
                alert('No data to download. Fetch data first.');
                return;
            }
            
            let ws = XLSX.utils.aoa_to_sheet(sheetData);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
            XLSX.writeFile(wb, "GoogleSheetData.xlsx");
        }
