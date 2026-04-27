/* venuedb.js - extracted application JS

This file contains the main behavior for loading the Excel file, building
the hierarchical country->city->venue UI, wiring filters, and collapse
behavior. Each top-level function is documented below.

Functions:
- loadExcelFile(): loads the Excel file via XHR and initializes the page data
- populateFilters(): populates the country and file type filter dropdowns
- displayData(data): renders the hierarchical UI (countries -> cities -> venues)
- applyFilters(): filters the loaded data and re-renders

Event wiring at bottom attaches filter change handlers and reset button.
*/

// CONFIGURATION: Change this to the path of your Excel file
const EXCEL_FILE_PATH = 'venue_list.xlsx'; // Update this path to your Excel file

let allData = []; // Store all data
let headers = []; // Store headers
let countryColumnIndex = -1; // Index of Country column
let fileTypeColumnIndex = -1; // Index of File Type column
let venueColumnIndex = -1; // Index of Venue column
let cityColumnIndex = -1; // Index of City column
// Behavior: whether clicking a country expands all cities (true) or keeps cities collapsed (false)
let expandCountryAll = false;

// Load the Excel file automatically on page load
window.addEventListener('load', function() {
    loadExcelFile();
    // initialize toggle from localStorage
    try {
        const stored = localStorage.getItem('expandCountryAll');
        expandCountryAll = stored === 'true' ? true : false;
    } catch (e) { expandCountryAll = false; }
    const toggle = document.getElementById('country_expand_toggle');
    if (toggle) {
        toggle.checked = expandCountryAll;
        toggle.addEventListener('change', function() {
            expandCountryAll = !!this.checked;
            try { localStorage.setItem('expandCountryAll', expandCountryAll ? 'true' : 'false'); } catch(e){}
        });
    }
});

/* loadExcelFile()
   Loads EXCEL_FILE_PATH via XHR as arraybuffer, parses with XLSX, and
   initializes `headers` and `allData`, then populates filters and renders.
*/
function loadExcelFile() {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', EXCEL_FILE_PATH, true);
    xhr.responseType = 'arraybuffer';

    xhr.onload = function(e) {
        if (this.status == 200) {
            // Get and display file last modified date
            var lastModified = xhr.getResponseHeader('Last-Modified');
            if (lastModified) {
                var date = new Date(lastModified);
                document.getElementById('file_info').innerHTML = ' Last Updated: <strong>' + date.toLocaleString() + '</strong></small>';
            } else {
                document.getElementById('file_info').innerHTML = '<small>File: <strong>' + EXCEL_FILE_PATH + '</strong></small>';
            }
            var data = new Uint8Array(xhr.response);
            var work_book = XLSX.read(data, {type:'array'});
            var sheet_name = work_book.SheetNames;
            var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {header:1});

            if(sheet_data.length > 0) {
                headers = sheet_data[0];
                allData = sheet_data;

                // Find column indexes for Country, File Type, and Venue
                countryColumnIndex = headers.findIndex(h => h && h.toString().toLowerCase().includes('country'));
                fileTypeColumnIndex = headers.findIndex(h => h && (h.toString().toLowerCase().includes('file type') || h.toString().toLowerCase().includes('filetype') || h.toString().toLowerCase().includes('type')));
                venueColumnIndex = headers.findIndex(h => h && h.toString().toLowerCase().includes('venue'));
                cityColumnIndex = headers.findIndex(h => h && h.toString().toLowerCase().includes('city'));

                // Populate filter dropdowns
                populateFilters();

                // Display initial data
                displayData(allData);
            }
        } else {
            document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Error loading Excel file. Make sure:<br>1. The file "' + EXCEL_FILE_PATH + '" exists in the same folder as this HTML file<br>2. The filename is correct<br>3. You are running this through a local web server (not file://)</div>';
        }
    };

    xhr.onerror = function() {
        document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Error loading Excel file. If opening this HTML directly (file://), you need to:<br>1. Run a local web server (e.g., Python: <code>python -m http.server</code>)<br>2. Or place both files in the same folder and access via http://localhost</div>';
    };

    xhr.send();
}

/* populateFilters()
   Reads `allData` and fills the country and file-type select controls.
*/
function populateFilters() {
    const countryFilter = document.getElementById('country_filter');
    const fileTypeFilter = document.getElementById('filetype_filter');

    // Get unique countries
    if (countryColumnIndex !== -1) {
        const countries = new Set();
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][countryColumnIndex]) {
                countries.add(allData[i][countryColumnIndex]);
            }
        }
        countries.forEach(country => {
            const option = document.createElement('option');
            option.value = country;
            option.textContent = country;
            countryFilter.appendChild(option);
        });
    } else {
        countryFilter.disabled = true;
        countryFilter.parentElement.innerHTML += '<small class="text-muted d-block">No "Country" column found</small>';
    }

    // Get unique file types (split multiple types in same cell)
    if (fileTypeColumnIndex !== -1) {
        const fileTypes = new Set();
        for (let i = 1; i < allData.length; i++) {
            if (allData[i][fileTypeColumnIndex]) {
                // Split by common delimiters: comma, slash, semicolon, pipe
                const cellValue = allData[i][fileTypeColumnIndex].toString();
                const types = cellValue.split(/[,\/;|]+/);
                types.forEach(type => {
                    const trimmedType = type.trim();
                    if (trimmedType) {
                        fileTypes.add(trimmedType);
                    }
                });
            }
        }
        // Sort file types alphabetically
        const sortedTypes = Array.from(fileTypes).sort();
        sortedTypes.forEach(fileType => {
            const option = document.createElement('option');
            option.value = fileType;
            option.textContent = fileType;
            fileTypeFilter.appendChild(option);
        });
    } else {
        fileTypeFilter.disabled = true;
        fileTypeFilter.parentElement.innerHTML += '<small class="text-muted d-block">No "File Type" column found</small>';
    }
}

/* displayData(data)
   Renders the provided `data` array (first row expected to be headers) into a
   hierarchical HTML structure inserted into `#excel_data`. This function builds
   a grouped object { country: { cities: { cityName: [venues] }}} and then
   creates collapsible buttons and tables for each group.
*/
function displayData(data) {
    if (data.length === 0) {
        document.getElementById('excel_data').innerHTML = '<div class="alert alert-info">No data to display</div>';
        document.getElementById('venue_count').textContent = 'Venues: 0';
        document.getElementById('country_count').textContent = 'Countries: 0';
        document.getElementById('city_count').textContent = 'Cities: 0';
        return;
    }

    // Count venues (rows with data, excluding header)
    const venueCount = data.length - 1;
    document.getElementById('venue_count').textContent = 'Venues: ' + venueCount;

    // Count unique countries
    const uniqueCountries = new Set();
    if (countryColumnIndex !== -1) {
        for (let i = 1; i < data.length; i++) {
            if (data[i][countryColumnIndex]) {
                uniqueCountries.add(data[i][countryColumnIndex]);
            }
        }
    }
    document.getElementById('country_count').textContent = 'Countries: ' + uniqueCountries.size;

    // Count unique cities
    const uniqueCities = new Set();
    if (cityColumnIndex !== -1) {
        for (let i = 1; i < data.length; i++) {
            if (data[i][cityColumnIndex]) {
                uniqueCities.add(data[i][cityColumnIndex]);
            }
        }
    }
    document.getElementById('city_count').textContent = 'Cities: ' + uniqueCities.size;

    // Build hierarchical grouping: Country -> City -> Venues
    function escapeHtml(str) {
        if (str === undefined || str === null) return '';
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    const rows = data.slice(1);
    const grouped = {};

    rows.forEach(r => {
        const country = (countryColumnIndex !== -1 && r[countryColumnIndex]) ? r[countryColumnIndex] : '(No country)';
        const city = (cityColumnIndex !== -1 && r[cityColumnIndex]) ? r[cityColumnIndex] : '(No city)';
        const venueName = (venueColumnIndex !== -1 && r[venueColumnIndex]) ? r[venueColumnIndex] : (r.join(' | ') || '(No venue)');

        if (!grouped[country]) grouped[country] = { cities: {}, venueCount: 0 };
        if (!grouped[country].cities[city]) grouped[country].cities[city] = [];
        grouped[country].cities[city].push({ name: venueName, row: r });
        grouped[country].venueCount += 1;
    });

    let output = '<ul class="country-list">';
    Object.keys(grouped).sort().forEach(country => {
        const cities = grouped[country].cities;
        const cityCount = Object.keys(cities).length;
        const venueCount = grouped[country].venueCount;

        output += '<li>';
        output += '<button class="collapsible"><span style="display:flex;align-items:center;"><span class="caret"></span><span style="margin-left:6px;font-weight:600">' + escapeHtml(country) + '</span></span><span class="badge-count">' + cityCount + ' cities • ' + venueCount + ' venues</span></button>';
        output += '<ul class="nested">';

        Object.keys(cities).sort().forEach(city => {
            const venues = cities[city];
            output += '<li>';
            // Place the city venue count badge in a right-aligned container so badges line up vertically
            // use an additional class so city buttons can have a fixed left width
            output += '<button class="collapsible city-collapsible"><span style="display:flex;align-items:center;" class="coll-left"><span class="caret"></span><span style="margin-left:6px">' + escapeHtml(city) + '</span></span><span class="coll-right"><span class="badge-count">' + venues.length + '</span></span></button>';
            output += '<ul class="nested">';

            // Render venues for this city as a compact table (grid style) so file types and other columns are visible
            if (venues.length > 0) {
                output += '<li>'; 
                output += '<div class="venue-table-container">';
                output += '<table class="table table-sm table-striped table-bordered">';
                // Determine visible columns (exclude Country and City)
                const visibleCols = headers.map((h, idx) => idx).filter(idx => idx !== countryColumnIndex && idx !== cityColumnIndex);
                // Compute equal column widths (percent) for visible columns
                const colCount = visibleCols.length || 1;
                const baseWidth = Math.floor(100 / colCount);
                const widths = visibleCols.map((cidx, i) => (i === colCount - 1) ? (100 - baseWidth * (colCount - 1)) : baseWidth);
                // Header row (only visible headers) with fixed widths
                output += '<thead><tr>' + visibleCols.map((idx, i) => '<th style="width:' + widths[i] + '%">' + escapeHtml(headers[idx] || '') + '</th>').join('') + '</tr></thead>';
                output += '<tbody>';
                venues.forEach(v => {
                    const row = v.row || [];
                    output += '<tr>' + visibleCols.map((idx, i) => '<td style="width:' + widths[i] + '%">' + escapeHtml(row[idx] || '') + '</td>').join('') + '</tr>';
                });
                output += '</tbody>';
                output += '</table>';
                output += '</div>';
                output += '</li>';
            } else {
                venues.forEach(v => {
                    output += '<li class="venue-item">' + escapeHtml(v.name) + '</li>';
                });
            }

            output += '</ul>';
            output += '</li>';
        });

        output += '</ul>';
        output += '</li>';
    });
    output += '</ul>';

    document.getElementById('excel_data').innerHTML = output;

    // Attach toggle behavior to collapsible buttons
    const collapsibles = document.querySelectorAll('#excel_data .collapsible');
    collapsibles.forEach(btn => {
        btn.addEventListener('click', function(e) {
            const caret = this.querySelector('.caret');
            const next = this.nextElementSibling;
            if (!next) return;

            const isOpen = next.style.display === 'block';

            // determine if this button is a country-level button
            let isCountryLevel = false;
            try {
                isCountryLevel = this.parentElement && this.parentElement.parentElement && this.parentElement.parentElement.classList && this.parentElement.parentElement.classList.contains('country-list');
            } catch (e) { isCountryLevel = false; }

            if (isOpen) {
                if (caret) caret.classList.remove('caret-down');
                // Close this nested list and any descendant nested lists
                const descendantNested = next.querySelectorAll('.nested');
                descendantNested.forEach(n => n.style.display = 'none');
                // Remove caret-down from any descendant carets
                const descendantCarets = next.querySelectorAll('.caret.caret-down');
                descendantCarets.forEach(c => c.classList.remove('caret-down'));
                next.style.display = 'none';
            } else {
                if (caret) caret.classList.add('caret-down');
                // Open this nested list (always show the immediate nested list)
                next.style.display = 'block';

                if (isCountryLevel) {
                    // Country-level click — follow toggle behavior
                    if (expandCountryAll) {
                        // expand all descendant nested lists and set carets
                        const descendantNested = next.querySelectorAll('.nested');
                        descendantNested.forEach(n => n.style.display = 'block');
                        const descendantCarets = next.querySelectorAll('.caret');
                        descendantCarets.forEach(c => c.classList.add('caret-down'));
                    } else {
                        // keep city-level nested lists collapsed; ensure descendant carets are cleared
                        const descendantNested = next.querySelectorAll('.nested');
                        descendantNested.forEach(n => n.style.display = 'none');
                        const descendantCarets = next.querySelectorAll('.caret.caret-down');
                        descendantCarets.forEach(c => c.classList.remove('caret-down'));
                        // keep only the country caret marked open (already added above)
                    }
                } else {
                    // city-level button: only toggle this city's nested content (we already opened `next`)
                    // nothing more to do here
                }
            }
        });
    });
}

/* applyFilters()
   Rebuilds a filtered data array based on the selected filters then calls
   `displayData()` to render the filtered view.
*/
function applyFilters() {
    const selectedCountry = document.getElementById('country_filter').value;
    const selectedFileType = document.getElementById('filetype_filter').value;

    let filteredData = [headers]; // Start with headers

    for (let i = 1; i < allData.length; i++) {
        let includeRow = true;

        // Filter by country
        if (selectedCountry && countryColumnIndex !== -1) {
            if (allData[i][countryColumnIndex] !== selectedCountry) {
                includeRow = false;
            }
        }

        // Filter by file type (substring match to handle multiple types in one cell)
        if (selectedFileType && fileTypeColumnIndex !== -1) {
            const cellValue = allData[i][fileTypeColumnIndex] ? allData[i][fileTypeColumnIndex].toString() : '';
            if (!cellValue.includes(selectedFileType)) {
                includeRow = false;
            }
        }

        if (includeRow) {
            filteredData.push(allData[i]);
        }
    }

    displayData(filteredData);
}

// Add event listeners for filters
document.getElementById('country_filter').addEventListener('change', applyFilters);
document.getElementById('filetype_filter').addEventListener('change', applyFilters);
document.getElementById('reset_filters').addEventListener('click', function() {
    document.getElementById('country_filter').value = '';
    document.getElementById('filetype_filter').value = '';
    displayData(allData);
});
