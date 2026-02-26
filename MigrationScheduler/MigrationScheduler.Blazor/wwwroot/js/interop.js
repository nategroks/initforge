/**
 * JavaScript interop functions for the Migration Scheduler Blazor application.
 * Provides file download and upload capabilities via the browser.
 */
window.interop = {
    /**
     * Downloads a file by creating a Blob and triggering a browser download.
     * @param {string} fileName - The default file name for the download.
     * @param {string} content - The file content as a string.
     * @param {string} mimeType - The MIME type (e.g., "application/xml", "application/json").
     */
    downloadFile: function (fileName, content, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        URL.revokeObjectURL(url);
    },

    /**
     * Opens a file picker dialog and returns the file content as a string.
     * @param {string} accept - File type filter (e.g., ".json", ".xml").
     * @returns {Promise<string>} The file content or empty string if cancelled.
     */
    uploadFile: function (accept) {
        return new Promise((resolve) => {
            const input = document.createElement("input");
            input.type = "file";
            input.accept = accept;
            input.style.display = "none";

            input.addEventListener("change", function () {
                if (input.files && input.files.length > 0) {
                    const reader = new FileReader();
                    reader.onload = function () {
                        resolve(reader.result);
                    };
                    reader.onerror = function () {
                        resolve("");
                    };
                    reader.readAsText(input.files[0]);
                } else {
                    resolve("");
                }
                document.body.removeChild(input);
            });

            input.addEventListener("cancel", function () {
                resolve("");
                document.body.removeChild(input);
            });

            document.body.appendChild(input);
            input.click();
        });
    }
};
