**PowerPoint Blob Consolidator**
This Azure Function consolidates PowerPoint presentations (PPTX) from a specified folder in Azure Blob Storage into a single PowerPoint file. It also extracts images from each slide in the PPTX files and includes them in the consolidated file. After consolidation, the original PowerPoint files are stored in a separate folder within the same container and then deleted.

**Features**
1. Retrieve PowerPoint files from a specific folder in an Azure Blob Storage container.
2. Consolidate multiple PowerPoint presentations into one, combining the slides from all the presentations.
3. Extract images from each slide and include them in the new consolidated presentation.
4. Store the consolidated PowerPoint file back into the same folder.
5. Move original PowerPoint files to a separate subfolder and delete them from their original location.
6. Supports folder-based file management (e.g., DEV, UAT, PROD).

**Prerequisites**
1. Azure subscription and an Azure Storage account.
2. Python 3.7+.
3. Azure Functions environment.
4. The following Python packages:
    a. azure-functions
    b. azure-storage-blob
    c. python-pptx
    d. Pillow
