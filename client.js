// Add this script to your website's HTML
<script src="https://cdn.tailwindcss.com"></script>

<!-- File selector and output container -->
<div class="max-w-4xl mx-auto mt-10 px-4">
    <div class="mb-6">
        <label class="block text-gray-700 text-sm font-bold mb-2" for="fileSelect">
            Select Excel File:
        </label>
        <select id="fileSelect" class="shadow border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline">
            <option value="Moodle_datein.xlsx">Moodle_datein.xlsx</option>
            <option value="PublicMoodleData.xlsx">PublicMoodleData.xlsx</option>
            <option value="PublicMoodleNewsfeed.xlsx">PublicMoodleNewsfeed.xlsx</option>
        </select>
    </div>
    <div id="jsonView"></div>
</div>

<script type="module">
    async function fetchExcelData(filename) {
        try {
            const response = await fetch(`https://excel-rest-8tn9c7y6p-akames-projects-01f01e48.vercel.app/api/excel-data?file=${encodeURIComponent(filename)}`);
            if (!response.ok) {
                throw new Error('Failed to fetch data');
            }
            const data = await response.json();
            displayData(data);
        } catch (error) {
            document.getElementById('jsonView').innerHTML = `
                <div class="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded">
                    Error: ${error.message}
                </div>
            `;
        }
    }

    function displayData(response) {
        if (!response.data || response.data.length === 0) {
            document.getElementById('jsonView').innerHTML = `
                <div class="bg-yellow-100 border border-yellow-400 text-yellow-700 px-4 py-3 rounded">
                    No data available
                </div>
            `;
            return;
        }

        const headers = response.headers;
        const headerRow = headers.map(header => `
            <th class="px-6 py-3 text-left font-semibold">${header}</th>
        `).join('');

        const rows = response.data.map(item => `
            <tr class="border-t hover:bg-gray-50">
                ${headers.map(header => `
                    <td class="px-6 py-4">${item[header] !== undefined ? item[header] : ''}</td>
                `).join('')}
            </tr>
        `).join('');

        document.getElementById('jsonView').innerHTML = `
            <div class="bg-white shadow-md rounded-lg overflow-hidden border border-gray-200">
                <div class="px-6 py-4 bg-gray-50 border-b border-gray-200">
                    <h2 class="text-lg font-semibold text-gray-800">
                        Data from ${response.filename}
                    </h2>
                </div>
                <div class="overflow-x-auto">
                    <table class="min-w-full text-sm text-gray-700">
                        <thead class="bg-gray-100">
                            <tr>${headerRow}</tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
            </div>
        `;
    }

    // Set up event listener for file selection
    document.getElementById('fileSelect').addEventListener('change', (e) => {
        fetchExcelData(e.target.value);
    });

    // Fetch initial data
    fetchExcelData('Moodle_datein.xlsx');
</script>