var currentDoc = app.activeDocument;
var firstPageNum = currentDoc.pages.length-1; // TO DO: Fazer a UI perguntar isso 
var currentPage = currentDoc.pages.item(firstPageNum);


var x = currentPage.marginPreferences.left;
var y = currentPage.marginPreferences.top;

importFiles(getFilesList(), firstPageNum);
linkFrames(firstPageNum)


// Abre a caixa de dialogo que seleciona os arquivos
function getFilesList(){
    return File.openDialog("Please choose a file", "*.doc;*.docx", true);
}

// Importa os arquivos e coloca cada um em uma página
function importFiles(files, pageNum){
    for(var i = 0; i < files.length; i++){
        if(pageNum + i > currentDoc.pages.length-1){
            currentDoc.pages.add();
        }
        var currentPage = currentDoc.pages.item(pageNum + i);
        currentPage.place(File(files[i]), [x,y], undefined, false,false)[0];
    }
}
// junta as text frames, mas não fazer o auto flow
function linkFrames(pageNum){
    var allPages = currentDoc.pages;
    var currentFrame = allPages[pageNum].textFrames[0];

    for(var i = 1; i < allPages.length; i++){
        if(pageNum + i >= allPages.length){
            break;
        } else{
            currentFrame.nextTextFrame = allPages[pageNum + i].textFrames[0];
            currentFrame = allPages[pageNum + i].textFrames[0];
        }
    }
}

/*



SCRIPT ANTIGO
//// Assuming Every Page has only one main textFrame ---


var myDocument = app.documents.item(0);
var firstPage = myDocument.pages.length-1;
var myPage = myDocument.pages.item(firstPage);


var x = myPage.marginPreferences.left;
var y = myPage.marginPreferences.top;
var path = "C:/Users/vinic/Documents/Editora/Digramações/João Benvindo (org) - Semiolinguistica/Arquivos originais/"
var filename = [
    "12 - Jaqueline Salviano corrigido.docx",
    "13 - Jesica Carvalho.docx",
    "14 - José Magno.docx",
    "15 - José Maria corrigido.docx",
    "16 - Luis Felipe atualizado em 27.03.docx",
    "17 - Maria Juliana corrigido.docx",
    "18 - Patrícia.docx",
    "Sobre os autores.docx"
]
for(var i = 0; i < filename.length; i++){
    if(i+firstPage > myDocument.pages.length-1){
        myDocument.pages.add();
    }
    var myPage = myDocument.pages.item(firstPage + i);
    var myStory = myPage.place(File(path+filename[i]), [x,y], undefined, false,false) [0];
}

var myDoc = app.documents[0];

var allPages = myDoc.pages;
var currentFrame = allPages[firstPage].textFrames[0];

for(var i = 1; i < allPages.length; i++){

    currentFrame.nextTextFrame = allPages[i+firstPage].textFrames[0];

    currentFrame = allPages[i+firstPage].textFrames[0];

    }
   */