// Esse script cria e formata os números de acordo com os capítulos do livro

// To do: mudar essa parte pra achar os títulos mais facilmente
var tittleStyle;
//tittleStyle = app.activeDocument.paragraphStyleGroups.item(1).paragraphStyles.item(0);
#include "showProps.jsx";
var xRef = app.documents[0].hyperlinks[0];
for (p in xRef) {
    $.writeln(p + ": " + xRef);
}

var currentDoc = app.activeDocument;
var selectedParagraphID = [];

paragraphStylesNames = currentDoc.paragraphStyles.everyItem().name
paragraphStylesID = currentDoc.paragraphStyles.everyItem().id;
for(i = 0; i < currentDoc.paragraphStyleGroups.length; i++){
    for(j = 0; j < currentDoc.paragraphStyleGroups.item(i).paragraphStyles.length; j++){
        paragraphStylesNames.push("("+ currentDoc.paragraphStyleGroups.item(i).name +") "+ currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).name)
        paragraphStylesID.push(currentDoc.paragraphStyleGroups.item(i).paragraphStyles.item(j).id)
    }
}

var myWindow = new Window ("dialog", "Gerar guias");

myWindow.orientation = "row";

var group1 = myWindow.add('group {alignment:["fill","fill"]}')
group1.orientation = "column";
var text1 = group1.add ("statictext", undefined, "Escolha o estilo de paragráfo:");
text1.alignment = "left"
var group1_1 = group1.add('group')
group1_1.orientation = "row";
var myDropdown = group1_1.add ("dropdownlist", undefined, paragraphStylesNames);
myDropdown.selection = 0;
var addButton = group1_1.add ('button {text: "Adicionar >>"}');

var group1_2 = group1.add('group {alignment:["left","fill"]}');
group1_2.orientation = "column";
var checkBox = group1_2.add ("checkbox", undefined, "Apagar marcações existentes");
checkBox.value = true;

var group1_3 = group1.add('group');
group1_3.orientation = "row";
okButton = group1_3.add("button", undefined, "OK");
okButton.enabled = false
group1_3.add("button", undefined, "Cancel");

var group2 = myWindow.add('group')
group2.orientation = "column";
var text2 = group2.add ("statictext", undefined, "Estilos de paragráfo selecionados:");
text2.alignment = "left"
var myList = group2.add ("listbox", undefined,);
myList.preferredSize.width = 240;
myList.preferredSize.height = 100;

addButton.onClick = function () {
    myList.add('item',myDropdown.selection.text);
    selectedParagraphID.push(paragraphStylesID[myDropdown.selection.index])
    okButton.enabled = true;
}

if(myWindow.show() == 1){
    tittleStyle = currentDoc.paragraphStyles.itemByID(selectedParagraphID[0])

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.appliedParagraphStyle = tittleStyle;
    tittleList = app.activeDocument.findText();
    numTittles = tittleList.length
    newContents = "";
    previusPage = 0;

    for (i = 0; i < numTittles; i++){
        currentPage = tittleList[i].parentTextFrames[0].parentPage.name
        if(currentPage !== previusPage){
            newContents = newContents + "<ViniSourceTag>\r"
        }
        previusPage = currentPage;
    }

    selection = app.selection[0]
    currentFrame = selection.parentTextFrames[0];
    currentFrame.contents = newContents


    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.findWhat = "<ViniSourceTag>"
    sourceList = app.activeDocument.findText();

    xRefFormat = app.activeDocument.crossReferenceFormats.item(8) // Formato "page number" (somente o número de página)

    previusPage = 0;
    sourceIndex = 0
    for(i = 0; i < numTittles; i++){
        currentPage = tittleList[i].parentTextFrames[0].parentPage.name
        if(currentPage !== previusPage){
            // ------ destination
            myHypTextDest = app.activeDocument.hyperlinkTextDestinations.add(tittleList[i]);
            myHypTextDest.name = tittleList[i].contents + i;

            // ------ source
            mySource = app.activeDocument.crossReferenceSources.add(sourceList[sourceIndex], xRefFormat);
            mySource.name = tittleList[i].contents + i;
            sourceIndex++;
            

            myHypDest = app.activeDocument.hyperlinkTextDestinations.itemByName(tittleList[i].contents + i);
            
            if (myHypDest.isValid) {
                myHyperLink = app.activeDocument.hyperlinks.add(mySource, myHypDest);
                myHyperLink.visible = false;
                myHyperLink.name = tittleList[i].contents + i;
            }
        }
        previusPage = currentPage;
    }
//DESCRIPTION: Exploring our one and only Cross Reference

   
}


//alert(tittleStyle.name)

