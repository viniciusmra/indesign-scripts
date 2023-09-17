/*
    Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: 2022
    Última atualização: abr/2023
    
	Descrição:
	Aplica estilo de objeto as imagens e estilo de tabela para as tabelas
*/




//@include showProps.jsx
var currentDoc = app.activeDocument;

var pages = currentDoc.pages.everyItem().getElements();

try{
    //showProps(currentDoc.stories.everyItem().tables.everyItem())
    //alert(pages[2].parent.parent.spreads[2].pages[1].name)
} 
catch(e){}

var images = currentDoc.allGraphics; // Cria uma lista com todas as imagens do livro
var tables = currentDoc.stories.everyItem().tables;

// Lista de estilos de parágrafos

objectStylesNames = currentDoc.objectStyles.everyItem().name;
objectStylesID = currentDoc.objectStyles.everyItem().id;
for(i = 0; i < currentDoc.objectStyleGroups.length; i++){
    for(j = 0; j < currentDoc.objectStyleGroups.item(i).objectStyles.length; j++){
        objectStylesNames.push("("+ currentDoc.objectStyleGroups.item(i).name +") "+ currentDoc.objectStyleGroups.item(i).objectStyles.item(j).name)
        objectStylesID.push(currentDoc.objectStyleGroups.item(i).objectStyles.item(j).id)
    }
}

tableStylesNames = currentDoc.tableStyles.everyItem().name;
tableStylesID = currentDoc.tableStyles.everyItem().id;
for(i = 0; i < currentDoc.tableStyleGroups.length; i++){
    for(j = 0; j < currentDoc.tableStyleGroups.item(i).tableStyles.length; j++){
        tableStylesNames.push("("+ currentDoc.tableStyleGroups.item(i).name +") "+ currentDoc.tableStyleGroups.item(i).tableStyles.item(j).name)
        tableStylesID.push(currentDoc.tableStyleGroups.item(i).tableStyles.item(j).id)
    }
}


var myObjectStyle;
var myTableStyle;

// Necessário para o "palette" funcionar (as barras "//" servem pro vscode não apontar erro)
//@targetengine "mysession"
var win = createDialog();
win.show();

function createDialog(){
    var myWindow = new Window ("palette", "Formatar imagens e tabelas");
    //myWindow.spacing = 0

    myWindow.add("statictext", undefined, "Número de imagens: " + images.length).alignment = "left";
    myWindow.add("statictext", undefined, "Número de tabelas: " + tables.length).alignment = "left";

    // Group 1
    var group1 = myWindow.add('group {alignment:["left","fill"]}');
    group1.orientation = "row";
    group1.add("statictext", undefined, "Estilos de imagem ");
    var myDropdown1 = group1.add("dropdownlist", undefined, objectStylesNames);
    myDropdown1.selection = objectStylesNames.length - 1;
    myDropdown1.preferredSize.width = 200;

    // Group 2
    var group2 = myWindow.add('group {alignment:["left","fill"]}');
    group2.orientation = "row";
    group2.add("statictext", undefined, "Estilos de tabela ").preferredSize.width = 107;
    var myDropdown2 = group2.add("dropdownlist", undefined, tableStylesNames);
    myDropdown2.selection = tableStylesNames.length - 1;
    myDropdown2.preferredSize.width = 200;

    // Group 3
    var group3 = myWindow.add('group {alignment:["left","fill"]}');
    group3.orientation = "row";
    //group2.add("statictext", undefined, "Estilos de tabela ").preferredSize.width = 107;
    formatImagesButton = group3.add("button", undefined, "Formatar imagens");
    formatTablesButton = group3.add("button", undefined, "Formatar tabelas");
    //var myDropdown2 = group2.add("dropdownlist", undefined, tableStylesNames);
    //myDropdown2.selection = tableStylesNames.length - 1;
    //myDropdown2.preferredSize.width = 200;
    myWindow.addEventListener('keydown', function(e){
        alert("safdsf")
		
		if(e.keyName == 'Enter'){
        }
			
	})

    formatImagesButton.onClick = function(){
        for(var i = 0; i < images.length; i++){
            images[i].parent.appliedObjectStyle = app.activeDocument.objectStyles.item(myDropdown1.selection.text);
            images[i].fit(FitOptions.FILL_PROPORTIONALLY); 
        }

        for(var i = 0; i < images.length; i++){
            imagePage = images[i].parentPage.name-1 // Encontra o numero da página em que a imagem está
        
            // Checa se a imagem está em uma página que contem um TextFrame
            if(pages[imagePage].textFrames.length > 0 ){
                // Pega os valores das coordenadas do frame de texto da página [y1, x1, y2, x2]
                frameLocation = pages[imagePage].textFrames[0].geometricBounds
                frameWidth = frameLocation[3] - frameLocation[1];
                frameHeight = frameLocation[2] - frameLocation[0]
                frameX = frameLocation[1];
                
                // Pega os valores das coordenadas do frame de texto da página [y1, x1, y2, x2]
                imageLocation = images[i].geometricBounds;
                imageHeight = imageLocation[2] - imageLocation[0];
                imageWidth =  imageLocation[3] - imageLocation[1];
                
                imageRatio = imageHeight/imageWidth;   // Calcula a proporção da imagem 
                
                // Se a imagem tiver a altura maior que a largura, a altura dela será limitada à largura
                // do TextFrame, evitando que uma imagem não caiba no frame
                /*
                if(imageRatio > 1){ 
                    imageNewHeight = frameWidth;
                    imageNewWidth = imageNewHeight/imageRatio;
                    imageX = (frameWidth - imageNewWidth)/2
                }
                else{ // Se ela tiver a altura menor que a largura, a largura sera a mesma do TextFrame
                    imageNewWidth = frameWidth
                    imageNewHeight = imageNewWidth*imageRatio // Calcula a nova altura da imagem baseada na proporção
                    imageX = frameX
                }*/
                imageNewHeight = Math.min(frameHeight/2, (frameWidth/2)*imageRatio)
                imageNewWidth = Math.min(frameWidth/2, (frameHeight/2)/imageRatio)
                imageX = (frameWidth - imageNewWidth)/2
                images[i].geometricBounds = [imageLocation[0], imageX, imageLocation[0] + imageNewHeight, imageX + imageNewWidth]; 
                images[i].fit(FitOptions.FRAME_TO_CONTENT); // Ajusta o Frame da imagem a ela mesma
            
            }
        }
    }

    formatTablesButton.onClick = function(){
        for(var i = 0; i < pages.length; i++){
            if(pages[i].textFrames.length != 0){
                frame = pages[i].textFrames[0]
                frameLocation = pages[i].textFrames[0].geometricBounds
                frameWidth = frameLocation[3] - frameLocation[1];
                //var tables = frame.tables.everyItem().getElements(); // Cria uma lista com todas as imagens do livro
                for(var j = 0; j < tables.length; j++){
                    tables[j].width = frameWidth;
                    if(tables[j].rows.length > 1){
                        if(tables[j].rows[0].rowType == RowTypes.BODY_ROW){
                            tables[j].rows[0].rowType = RowTypes.HEADER_ROW;
                        }
                        tables[j].appliedTableStyle = app.activeDocument.tableStyles.item(myDropdown2.selection.text);
                        tables[j].clearTableStyleOverrides()
                        app.activeDocument.stories.everyItem().tables.everyItem().cells.everyItem().clearCellStyleOverrides(true);
                    }
        
                }
            }
        }
    }

    return myWindow;
}

/*
var images = app.activeDocument.allGraphics; // Cria uma lista com todas as imagens do livro


imagesNumber = app.activeDocument.allGraphics.length;

//Reseta a propoção de todas as imagens do livro e atribui um ObjectStyle
var myObjectStyle = prompt("Nome do ObjectStyle que será aplicado em todas as imagens", "anchor");

for(var i = 0; i < images.length; i++){
    images[i].parent.appliedObjectStyle = app.activeDocument.objectStyles.item(myObjectStyle);
    images[i].fit(FitOptions.FILL_PROPORTIONALLY); 
}

for(var i = 0; i < imagesNumber; i++){
    imagePage = images[i].parentPage.name // Econtra o numero da página em que a imagem está

    // Checa se a imagem está em uma página que contem um TextFrame
    if(frameLocation = app.activeDocument.pages.item(imagePage).textFrames.length > 0 ){
        // Pega os valores das coordenadas do frame de texto da página [y1, x1, y2, x2]
        frameLocation = app.activeDocument.pages.item(imagePage).textFrames[0].geometricBounds
        frameWidth = frameLocation[3] - frameLocation[1];
        frameX = frameLocation[1];
        
        // Pega os valores das coordenadas do frame de texto da página [y1, x1, y2, x2]
        imageLocation = images[i].geometricBounds;
        imageHeight = imageLocation[2] - imageLocation[0];
        imageWidth =  imageLocation[3] - imageLocation[1];
        
        imageRatio = imageHeight/imageWidth;   // Calcula a proporção da imagem 
        
        // Se a imagem tiver a altura maior que a largura, a altura dela será limitada à largura
        // do TextFrame, evitando que uma imagem não caiba no frame
        if(imageRatio > 1){ 
            imageNewHeight = frameWidth;
            imageNewWidth = imageNewHeight/imageRatio;
            imageX = (frameWidth - imageNewWidth)/2
        }
        else{ // Se ela tiver a altura menor que a largura, a largura sera a mesma do TextFrame
            imageNewWidth = frameWidth
            imageNewHeight = imageNewWidth*imageRatio // Calcula a nova altura da imagem baseada na proporção
            imageX = frameX
        }
        images[i].geometricBounds = [imageLocation[0], imageX, imageLocation[0] + imageNewHeight, imageX + imageNewWidth]; 
        images[i].fit(FitOptions.FRAME_TO_CONTENT); // Ajusta o Frame da imagem a ela mesma
    
    }
}
alert("Images: " + imagesNumber)
*/