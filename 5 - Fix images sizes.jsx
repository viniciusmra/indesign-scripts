/*
Script para consertar o tamanho e a proporção de todas as imagens do texto
Vinícius Alves - viniciusmra@gmail.com - agosto/2021

*/
var images = app.activeDocument.allGraphics; // Cria uma lista com todas as imagens do livro


imagesNumber = app.activeDocument.allGraphics.length;

//Reseta a propoção de todas as imagens do livro e atribui um ObjectStyle
var myObjectStyle = prompt("Nome do ObjectStyle que será aplicado em todas as imagens", "anchor");

for(var i = 0; i < imagesNumber; i++){
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
