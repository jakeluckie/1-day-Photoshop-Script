var doc = app.activeDocument;
const filePath = doc.path.toString();
const rootPath = filePath.slice(0, -8); // Slices off the end of the path string " - Working"

const pCodePatt = /[A-Z0-9]{9}/; // Regular Expression for matching 1-day product codes. (9 characters of A to Z or 0 to 9)
const groupLayerPatt = /^[_][0-9]$|^[_][1][0-9]$/; // RegEx for matching group names exactly (from _1 to _19) allowing the max of 20 zoom images

const workingFolder = new Folder(rootPath + '/Working'); // Current working folder where the .psd files live
const saveForWebFolder = new Folder(rootPath + '/Save for Web');  // Folder where the .jpg images are saved to on the photohraphy drive
const imageUploadsFolder = new Folder('/Data/!~!Image%20Upload!~!/1-day'); // Folder on S drive where the .jpg images are saved to for uploading to the site.
const accredoUploadsFolder = new Folder('/Data/!~!Image%20Upload!~!/1-day/__AccredoImages');
const filePathSplit = filePath.split("%20"); // Splits file path string into separate array items, %20 or a space, being the divider
const codeAndFolder = filePathSplit[filePathSplit.length-1]; // The last item in the array "filePathSplit" 
const codeAndFolderSplit = codeAndFolder.split("/"); // Splits the product code from the working folder

const productCode = codeAndFolderSplit[0]; // final product code extracted from filePath
const folder = codeAndFolderSplit[1]; // Working folder extracted from filePath

var productColour = 'NOTSET'; //sets default colour to be NOTSET
var savedState = app.activeDocument.activeHistoryState

var firstZoom = 0;

var win = new Window("window{text:'Saving Files...',bounds:[100,100,400,150],bar:Progressbar{bounds:[20,20,280,31] , value:0,maxvalue:100}};");
win.show();
var progressBarValue = 0;

// Function to save a file
// filePath is passed in when the function is called. pointing to the directory where the file will be saved.
// JPEG compression changed with saveForWebSettings.quality (1-100)
function saveFile (filePath) {
    const saveLocation = new File(filePath); // converts filePath to of type File. exportDocument requires this type.
    const saveForWebSettings = new ExportOptionsSaveForWeb;
    saveForWebSettings.quality = 85;
    saveForWebSettings.format = SaveDocumentType.JPEG;
    saveForWebSettings.includeProfile = true;
    try {
        doc.exportDocument(saveLocation, ExportType.SAVEFORWEB, saveForWebSettings);
    } catch (e) {
        alert('ERROR! File Path Error. Please check your folder structure is correct OR check that you are connected to the Data drive. \nThank you, have a lovely day.');
    }
}

// Function to make sure all layers within the group are visible before saving
function changeLayerVisibility(groupLayer, isVisible) {
    if (isVisible) {
        groupLayer.visible = isVisible;

        for (j = 0; j < groupLayer.artLayers.length; j++) { // then cycle through layers in the group
            groupLayer.artLayers[j].visible = isVisible; // layers are set to visible / not visible through setting the parameter isVisible being passed to the function
        }
    } else {
        groupLayer.visible = isVisible;
    }
}

function resizeToSmallAndSaveFile() {
    doc.resizeImage(150)
    saveFile(accredoUploadsFolder + '/' + productCode + '_small.jpg');
    doc.activeHistoryState = savedState;
}

function setProductColour() {
    if (doc.artLayers.length !== 0){
        for (j = 0; j < doc.artLayers.length; j++){
            if(!doc.artLayers[j].isBackgroundLayer) { // if the layer is not the background layer
                // check whether the layer is a solid shape layer, has a fill opacity of 0, and the blend mode is color burn. a hacky way of checking
                // the right layer to read the colour value from
                if(doc.artLayers[j].kind === LayerKind.SOLIDFILL && doc.artLayers[j].fillOpacity === 0 && doc.artLayers[j].blendMode === BlendMode.COLORBURN) {
                    productColour = doc.artLayers[j].name;
                    productColour = productColour.toUpperCase();
                }
            }
        }
    }
}

function getNumberOfZooms() {
    var numberOfZooms = 0;
    for (k = 0; k < doc.layerSets.length; k++) {
        if (doc.layerSets[k].artLayers.length !== 0) {
            var numberOfZooms = numberOfZooms + 1;
    } 
    return numberOfZooms;
}

function cycleThroughLayers() {
    setProductColour();
    var zoomNumber = getNumberOfZooms();

    for (h = 0; h < doc.layerSets.length; h++){
        doc.layerSets[h].visible = false;
    }
    for (i = 0; i < doc.layerSets.length; i++) { // Cycles through groups in photoshop
        if (doc.layerSets[i].artLayers.length !== 0) { // If there are layers in the group

            var groupLayer = doc.layerSets[i];
            var groupName = groupLayer.name;

            changeLayerVisibility(groupLayer, true);

            if (groupName == '_1') {
                firstZoom = groupLayer // sets firstZoom to equal the group '_1'

                saveFile(accredoUploadsFolder + '/' + productCode + '_zoom.jpg'); // saves images for accredo
            }; 

            saveFile(saveForWebFolder + '/' + productCode + '_' + productColour + groupName + '.jpg');
            saveFile(imageUploadsFolder + '/' + productCode + '_' + productColour + groupName + '.jpg');

            changeLayerVisibility(groupLayer, false);

            progressBarValue = progressBarValue + 100 / zoomNumber;
            win.bar.value = progressBarValue;
        } 
    }
    changeLayerVisibility(firstZoom, true); // makes _1 layer visible before saving (so the preview is visible in Adobe Bridge)
    resizeToSmallAndSaveFile();

    doc.save();
    win.close();
}

function main () {
    if (doc.name == '_zoom.psd') {
        if (productCode.length == 9 && pCodePatt.test(productCode)) { // checks if the product code on the folder is the right length & if it matches the RegEx rule of pCodePatt
            if (doc.name == '_zoom.psd') {
                cycleThroughLayers();
            } 
        } else {
            alert('ERROR! Something\'s wrong with the code on the folder. Please check the folder name. \nNote: product codes on folders must be 9 digits long.');
        }
    } else {
        alert('ERROR! Please make sure you are using a _zoom.psd document. \nThank you, have a lovely day.');
    }
}

main();

