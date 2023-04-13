docPages = app.activeDocument.pages;
for(var i = 0; i < docPages.length; i++){
    if(docPages.item(i).textFrames.length != 0){
        frame = docPages.item(i).textFrames[0]
        frameLocation = docPages.item(i).textFrames[0].geometricBounds
        frameWidth = frameLocation[3] - frameLocation[1];
        var tables = frame.tables.everyItem().getElements(); // Cria uma lista com todas as imagens do livro
        for(var j = 0; j < tables.length; j++){
            tables[j].width = frameWidth;
            if(tables[j].rows.length > 1){
                if(tables[j].rows[0].rowType == RowTypes.BODY_ROW){
                    alert(i)
                    tables[j].rows[0].rowType = RowTypes.HEADER_ROW;
                }
                tables[j].appliedTableStyle = app.activeDocument.tableStyles.item("Table");
                tables[j].clearTableStyleOverrides()
                app.activeDocument.stories.everyItem().tables.everyItem().cells.everyItem().clearCellStyleOverrides(true);
            }

        }
    }
}

/*
for (s=0; s<app.activeDocument.stories.length; s++)  
for (t=0; t<app.activeDocument.stories[s].tables.length; t++)  
{  
     app.activeDocument.stories[s].tables[t].appliedTableStyle = "Table";  
     app.activeDocument.stories[s].tables[t].clearTableStyleOverrides();
}  

// Clear All Overrides

var allStories = app.activeDocument.stories.everyItem();
// Remove overrides from all footnotes
try{
   allStories.footnotes.everyItem().texts.everyItem().clearOverrides(myOverrideType);
}
catch (e){alert ("No footnotes!")}

// Remove overrides from all table
try{
   allStories.tables.everyItem().cells.everyItem().paragraphs.everyItem().clearOverrides(myOverrideType);
}
catch (e){alert ("No tables!1")}

// Remove overrides from all cells
try{
allStories.tables.everyItem().cells.everyItem().clearCellStyleOverrides(true);
}
catch (e){alert ("No tables!2")}

alert("Overrides cleared!"); 
*/