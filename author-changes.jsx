//@include showProps.jsx
var currentDoc = app.activeDocument;
var spreads = currentDoc.spreads;

fileContent = "";
var matchFinds = [];
matchIndex = 0;
//alert(file.read().split("$%"));


//@targetengine "session"


//var window = new Window("palette", "title");
//var result = window.show();



var myWindow = new Window ('palette', "Correï¿½ï¿½es");

var group1 = myWindow.add('group {alignment:["fill","fill"]}')
group1.orientation = "row";
var selectButton = group1.add ('button {text: "Selecionar"}');

var changeList = myWindow.add ("listbox", undefined, "",{numberOfColumns: 3, showHeaders: true,columnTitles: ["#", "Página", "Onde"]});
changeList.preferredSize.width = 500;
changeList.preferredSize.height = 300;

var group2 = myWindow.add('group {alignment:["left","fill"]}')
group2.orientation = "row";
var group2_1 = group2.add('panel', [0,0,250,250], 'Como está:', {borderStyle:'black'})
group2_1.orientation = "column";
var oldText = group2_1.add('statictext', [0,0,200,200], '', {multiline: true});
oldText.alignment = 'left'

var group2_2 = group2.add('panel', [250,0,500,250], 'Como deveria ser:', {borderStyle:'black'})
group2_2.orientation = "column";
var newText = group2_2.add('statictext', [0,0,200,200], '', {multiline: true});

var group3 = myWindow.add('group {alignment:["left","fill"]}')
var changeButton = group3.add ('button {text: "Alterar"}');
var previusButton = group3.add ('button {text: "<"}');
var nextButton = group3.add ('button {text: ">"}');

changeButton.enabled = false;
previusButton.enabled = false;
nextButton.enabled = false;

changeList.onChange = function(){
    oldText.text = data[changeList.selection][2];
    newText.text = data[changeList.selection][3];
    searchConfig();
    findText(data[changeList.selection][2], data[changeList.selection][0]);
    changeButton.enabled = true;
}

selectButton.onClick = function(){
    data = [""];
    var file = File.openDialog();
    file.open("r");
    fileContent = file.read();
    fileLines = fileContent.split("\n");
    for(var i = 1; i < fileLines.length; i++){
        tempData = [];
        fileCells = fileLines[i].split("$%");
        if(fileCells.length == 4){
            with (changeList.add ("item", i)){
                subItems[0].text = fileCells[0];
                subItems[1].text = fileCells[1];
                }
            tempData.push(fileCells[0], fileCells[1], fileCells[2], fileCells[3])
        }
        data.push(tempData)
    }
}

changeButton.onClick = function () {
    //alert(currentDoc.selection[0].text)
    currentDoc.selection[0].contents = data[changeList.selection][3]
}

previusButton.onClick = function () {
    if(matchIndex > 0){
        matchIndex--
        currentDoc.selection = matchFinds[matchIndex]
    }
}

nextButton.onClick = function () {
    if(matchIndex < matchFinds.length-1){
        matchIndex++;
        currentDoc.selection = matchFinds[matchIndex]
    }
}

myWindow.show();


function searchConfig(){
    //Limpa as preferências do Grep
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    //Configura a busca
    app.findChangeGrepOptions.includeFootnotes = true;
    app.findChangeGrepOptions.includeHiddenLayers = true;
    app.findChangeGrepOptions.includeLockedLayersForFind = false;
    app.findChangeGrepOptions.includeLockedStoriesForFind = false;
    app.findChangeGrepOptions.includeMasterPages = false;

    //Limpa as preferências do texto
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    //Configura a busca
    app.findChangeTextOptions.includeFootnotes = true;
}

function findText(textToFind, page){
    matchFinds = [];
    previusButton.enabled = false;
    nextButton.enabled = false;
    if(page != ""){
        page = parseInt(page);
        var sp = findSpread(page-1);
        sp.select();
    }
    
    app.findTextPreferences.findWhat = textToFind;
    var finds = app.activeDocument.findText();

    for(var i = 0; i < finds.length; i++){
        if(finds[i].parentTextFrames[0].parentPage.name == page){
           matchFinds.push(finds[i]);
        }
    }
    if(matchFinds.length > 1){
        alert("UEPA: " + matchFinds.length)
        previusButton.enabled = true;
        nextButton.enabled = true;
    }
    matchIndex = 0;
    currentDoc.selection = matchFinds[matchIndex]
}

function findSpread(index) {
    for(var i = 0; i < spreads.length; i++){
        for(var j = 0; j < spreads[i].pages.length; j++){
            if(index === spreads[i].pages[j].name-1){
                return spreads[i];
            }
        }
    }
}
