currentDoc = app.activeDocument;
selection = currentDoc.selection[0]
selectionPage = currentDoc.selection[0].parentTextFrames[0].parentPage.name

xRefFormat = app.activeDocument.crossReferenceFormats.item(8)

word = selection.contents;
if(word !== ""){
    destinationList = findText(word);
    //alert(destinationList)
    sourceTags = word;
    previusPage = 0
    for (i = 0; i < destinationList.length; i++){
        currentPage = destinationList[i].parentTextFrames[0].parentPage.name
        if(currentPage !== previusPage && currentPage !== selectionPage && currentPage >= 7){
            sourceTags = sourceTags + " <ViniIndiceTag>,"
        }
        previusPage = currentPage;
    }
    sourceTags = sourceTags.slice(0,sourceTags.length-1);
    sourceTags = sourceTags + ";"
    selection.contents = sourceTags;

    /*app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.findWhat = "<ViniIndiceTag>,"
    sourceList = app.activeDocument.findText();*/   
    sourceList = findText("<ViniIndiceTag>,");
    charStyle(sourceList);
    sourceList = findText("<ViniIndiceTag>;");
    charStyle(sourceList);
    sourceList = findText("<ViniIndiceTag>");




    previusPage = 0;
    sourceIndex = 0
    removeHyperlinkAnchor(word);
    for(i = 0; i < destinationList.length; i++){
        currentPage = destinationList[i].parentTextFrames[0].parentPage.name
        if(currentPage !== previusPage && currentPage !== selectionPage && currentPage >= 7){
            // ------ destination
            myHypTextDest = currentDoc.hyperlinkTextDestinations.add(destinationList[i]);
            myHypTextDest.name = "IR_" + word + "_" + currentPage;

            // ------ source
            mySource = currentDoc.crossReferenceSources.add(sourceList[sourceIndex], xRefFormat);
            mySource.name = "IR_" + word + "_" + currentPage;
            sourceIndex++;

            myHypDest = currentDoc.hyperlinkTextDestinations.itemByName("IR_" + word + "_" + currentPage);
            if (myHypDest.isValid) {
                myHyperLink = currentDoc.hyperlinks.add(mySource, myHypDest);
                myHyperLink.visible = true;
                myHyperLink.name = "IR_" + word + "_" + currentPage;
            }
        }
        previusPage = currentPage;
    }
}

function charStyle(list){
    for(i = 0; i < list.length; i++){
        list[i].appliedCharacterStyle = "indiceRemissivo";
    }
}

function findText(text){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.findWhat = text
    return app.activeDocument.findText();
}

//alert(currentDoc.hyperlinkTextDestinations.item(0).name)
function removeHyperlinkAnchor(tag){
    for(i = currentDoc.hyperlinkTextDestinations.everyItem().getElements().length-1; i > -1; i--){
        if(currentDoc.hyperlinkTextDestinations.item(i).name.substring(0,3 + tag.length) == "IR_" + tag){
            currentDoc.hyperlinkTextDestinations.item(i).remove();
        }
    }
}

/*
des = currentDoc.hyperlinkTextDestinations.everyItem().getElements().length + "\r"
for(i = 0; i < currentDoc.hyperlinkTextDestinations.everyItem().getElements().length; i++){
    des = des + currentDoc.hyperlinkTextDestinations.everyItem().getElements()[i].name + "\r"
}
alert(des)*/
