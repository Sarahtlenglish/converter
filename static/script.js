document.addEventListener('DOMContentLoaded', function() {
    const excelToXmlNav = document.getElementById('excelToXmlNav');
    const xmlToExcelCsvNav = document.getElementById('xmlToExcelCsvNav');
    const excelToCsvNav = document.getElementById('excelToCsvNav');
    const csvImageDownloaderNav = document.getElementById('csvImageDownloaderNav');
    const excelToXmlContent = document.getElementById('excelToXmlContent');
    const xmlToExcelCsvContent = document.getElementById('xmlToExcelCsvContent');
    const excelToCsvContent = document.getElementById('excelToCsvContent');
    const csvImageDownloaderContent = document.getElementById('csvImageDownloaderContent');
    const converterTypeInput = document.createElement('input');

    const converterHeading = document.getElementById('converterHeading');
    const converterDescription = document.getElementById('converterDescription');
    const skabelonLink = document.getElementById('skabelonLink');

    converterTypeInput.type = 'hidden';
    converterTypeInput.name = 'converter_type';
    document.getElementById('uploadFormExcelToXml').appendChild(converterTypeInput);

    excelToXmlNav.addEventListener('click', function() {
        excelToXmlContent.style.display = 'block';
        xmlToExcelCsvContent.style.display = 'none';
        excelToCsvContent.style.display = 'none';
        csvImageDownloaderContent.style.display = 'none';
        setActiveNav(excelToXmlNav);
        converterTypeInput.value = 'excel_to_xml';
        updateConverterInfo('Excel to XML', 'Upload your Excel file and convert it to XML format.', true);
    });

    xmlToExcelCsvNav.addEventListener('click', function() {
        excelToXmlContent.style.display = 'none';
        xmlToExcelCsvContent.style.display = 'block';
        excelToCsvContent.style.display = 'none';
        csvImageDownloaderContent.style.display = 'none';
        setActiveNav(xmlToExcelCsvNav);
        converterTypeInput.value = 'xml_to_excel_csv';
        updateConverterInfo('XML to Excel/CSV', 'Upload your XML file and convert it to both Excel and CSV formats.', false);
    });

    excelToCsvNav.addEventListener('click', function() {
        excelToXmlContent.style.display = 'none';
        xmlToExcelCsvContent.style.display = 'none';
        excelToCsvContent.style.display = 'block';
        csvImageDownloaderContent.style.display = 'none';
        setActiveNav(excelToCsvNav);
        converterTypeInput.value = 'excel_to_csv';
        updateConverterInfo('Excel to CSV', 'Upload your Excel file and convert it to CSV format.', false);
    });

    csvImageDownloaderNav.addEventListener('click', function() {
        excelToXmlContent.style.display = 'none';
        xmlToExcelCsvContent.style.display = 'none';
        excelToCsvContent.style.display = 'none';
        csvImageDownloaderContent.style.display = 'block';
        setActiveNav(csvImageDownloaderNav);
        converterTypeInput.value = 'csv_image_downloader';
        updateConverterInfo('CSV Image Downloader', 'Upload your CSV file to download all images.', false);
    });

    function setActiveNav(activeElement) {
        excelToXmlNav.classList.remove('active');
        xmlToExcelCsvNav.classList.remove('active');
        excelToCsvNav.classList.remove('active');
        csvImageDownloaderNav.classList.remove('active');
        activeElement.classList.add('active');
    }

    function updateConverterInfo(heading, description, showSkabelon) {
        converterHeading.textContent = heading;
        converterDescription.textContent = description;
        skabelonLink.style.display = showSkabelon ? 'block' : 'none';
    }

    initializeForm('uploadFormExcelToXml', 'fileInputExcelToXml', 'fileInfoExcelToXml', 'feedbackExcelToXml', 'progressTextExcelToXml', 'application/xml', 'converted_file.xml');
    initializeForm('uploadFormXmlToExcelCsv', 'fileInputXmlToExcelCsv', 'fileInfoXmlToExcelCsv', 'feedbackXmlToExcelCsv', 'progressTextXmlToExcelCsv', 'application/zip', 'converted_files.zip');
    initializeForm('uploadFormExcelToCsv', 'fileInputExcelToCsv', 'fileInfoExcelToCsv', 'feedbackExcelToCsv', 'progressTextExcelToCsv', 'text/csv', 'converted_file.csv');
    initializeForm('uploadFormCsvImageDownloader', 'fileInputCsvImageDownloader', 'fileInfoCsvImageDownloader', 'feedbackCsvImageDownloader', 'progressTextCsvImageDownloader', 'application/zip', 'product_images.zip');

    function initializeForm(formId, fileInputId, fileInfoId, feedbackId, progressTextId, mimeType, downloadFilename) {
        const form = document.getElementById(formId);
        const fileInput = document.getElementById(fileInputId);
        const fileDropArea = form.querySelector('.file-drop-area');
        const fileInfo = document.getElementById(fileInfoId);

        // Prevent default behavior for drag-and-drop events
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            fileDropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        // Highlight file drop area on drag enter/over
        ['dragenter', 'dragover'].forEach(eventName => {
            fileDropArea.addEventListener(eventName, () => fileDropArea.classList.add('highlight'), false);
        });

        // Remove highlight on drag leave/drop
        ['dragleave', 'drop'].forEach(eventName => {
            fileDropArea.addEventListener(eventName, () => fileDropArea.classList.remove('highlight'), false);
        });

        // Handle drop event
        fileDropArea.addEventListener('drop', function(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            fileInput.files = files;
            updateFileInfo();
        });

        // Handle file input change
        fileInput.addEventListener('change', function() {
            updateFileInfo();
        });

        // Update file info on drop or file selection
        function updateFileInfo() {
            const files = fileInput.files;
            if (files.length > 0) {
                fileInfo.innerHTML = ''; // Clear existing content
                
                // Create file list container
                const fileList = document.createElement('div');
                fileList.className = 'file-list';
                
                // Add each file as an entry
                Array.from(files).forEach((file, index) => {
                    const fileEntry = document.createElement('div');
                    fileEntry.className = 'file-entry';
                    
                    const fileName = document.createElement('span');
                    fileName.textContent = file.name;
                    fileName.className = 'file-name';
                    
                    const removeButton = document.createElement('button');
                    removeButton.textContent = 'Fjern';
                    removeButton.type = 'button';
                    removeButton.className = 'remove-file-btn';
                    removeButton.onclick = () => removeFile(index);
                    
                    fileEntry.appendChild(fileName);
                    fileEntry.appendChild(removeButton);
                    fileList.appendChild(fileEntry);
                });
                
                fileInfo.appendChild(fileList);
            } else {
                fileInfo.textContent = '';
            }
        }

        function removeFile(indexToRemove) {
            const dt = new DataTransfer();
            const files = fileInput.files;
            
            for (let i = 0; i < files.length; i++) {
                if (i !== indexToRemove) {
                    dt.items.add(files[i]);
                }
            }
            
            fileInput.files = dt.files;
            updateFileInfo();
        }

        // Handle form submission
        form.addEventListener('submit', function(event) {
            event.preventDefault();
            if (fileInput.files.length === 0) {
                document.getElementById(feedbackId).textContent = 'Vælg venligst en fil.';
                document.getElementById(feedbackId).style.color = 'red';
                return;
            }

            handleFileUpload(fileInput.files, feedbackId, progressTextId, mimeType, downloadFilename);
        });
    }

    function handleFileUpload(files, feedbackId, progressTextId, mimeType, downloadFilename) {
        const formData = new FormData();
        const converterType = document.querySelector('input[name="converter_type"]').value;
        
        // Handle files differently based on converter type
        if (converterType === 'xml_to_excel_csv') {
            // For xml_to_excel_csv, append all files
            Array.from(files).forEach(file => {
                formData.append('file', file);
            });
        } else {
            // For other converters, only use the first file
            formData.append('file', files[0]);
        }
        
        formData.append('converter_type', converterType);

        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/', true);
        xhr.responseType = 'blob';

        xhr.onload = function() {
            if (xhr.status === 200) {
                const blob = new Blob([xhr.response], { type: mimeType });
                const downloadLink = document.createElement('a');
                downloadLink.href = window.URL.createObjectURL(blob);
                
                // For xml_to_excel_csv, keep original downloadFilename (zip)
                if (converterType === 'xml_to_excel_csv') {
                    downloadLink.download = 'avis-data.zip';
                } else {
                    // For single file converters, use input filename
                    const fileExtension = downloadFilename.split('.').pop();
                    const baseInputName = files[0].name.split('.')[0];
                    downloadLink.download = `${baseInputName}.${fileExtension}`;
                }
                
                downloadLink.click();
                document.getElementById(feedbackId).textContent = 'Konvertering fuldført!';
                document.getElementById(feedbackId).style.color = '#2C3E50';
            } else {
                document.getElementById(feedbackId).textContent = 'En fejl opstod under upload.';
                document.getElementById(feedbackId).style.color = 'red';
            }
        };

        xhr.upload.onprogress = function(event) {
            const percent = (event.loaded / event.total) * 100;
            const progressElement = document.getElementById(progressTextId);
            progressElement.textContent = `Progress: ${Math.round(percent)}%`;
            const progressBar = progressElement.previousElementSibling.querySelector('.progress');
            progressBar.style.width = `${percent}%`;
        };

        xhr.send(formData);
    }

    // Default to "Excel to XML" on page load
    converterTypeInput.value = 'excel_to_xml';
    excelToXmlContent.style.display = 'block';
    xmlToExcelCsvContent.style.display = 'none';
    excelToCsvContent.style.display = 'none';
    csvImageDownloaderContent.style.display = 'none';
    setActiveNav(excelToXmlNav);
});