document.getElementById('upload-form').onsubmit = async function(event) {
    event.preventDefault();

    const formData = new FormData(this);
    const response = await fetch('/merge', {
        method: 'POST',
        body: formData
    });

    const outputContainer = document.getElementById('output-container');
    if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'merged_output.xlsx';
        link.textContent = 'Download Merged File';
        outputContainer.innerHTML = '';
        outputContainer.appendChild(link);
    } else {
        const errorText = await response.text();
        outputContainer.textContent = `Error: ${errorText}`;
    }
};
