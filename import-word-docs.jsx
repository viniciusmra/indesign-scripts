/*
    Autor: Vinicius Alves - viniciusmra@gmail.com
    Data: 2021
    Descrição: Este script 


    TO-DO:
    - Quando os arquivos importados forem colocados no meio de uma sequência linkada, tem que haver uma forma de quebrar esse link
    - dar a opção de autoflow
    - organizar a interface
    - comentar o código

*/


//@include showProps.jsx;


var currentDoc = app.activeDocument;

var firstPageNum = currentDoc.pages.length-1; // TO DO: Fazer a UI perguntar isso 
var nPages = currentDoc.pages.length;
var pages = currentDoc.pages.everyItem().getElements();
var pageList = currentDoc.pages.everyItem().name;
var currentPage = currentDoc.pages.item(firstPageNum);
var x = currentPage.marginPreferences.left;
var y = currentPage.marginPreferences.top;
var files;

//alert(pageList[0])

//showProps(currentDoc.pages.everyItem())
// Cria a janela da interface
var myWindow = new Window ("dialog", "Importar arquivos");

//
var text1 = myWindow.add("statictext", undefined, "Selecione os arquivos que deseja importar");
var selectButton = myWindow.add ('button {text: "Selecionar"}');

var myDropdown = myWindow.add ("dropdownlist", undefined, pageList);
myDropdown.selection = nPages-1;

var fileListBox = myWindow.add("listbox", undefined);
fileListBox.preferredSize.width = 540;
fileListBox.preferredSize.height = 300;

var moveUpButton = myWindow.add ('button {text: "Mover para cima"}');
var moveDownButton = myWindow.add ('button {text: "Mover para baixo"}');
var delButton = myWindow.add ('button {text: "Remover arquivo"}');
moveDownButton.enabled = false;
moveUpButton.enabled = false;
delButton.enabled = false;

var okButton = myWindow.add ("button", undefined, "Importar", {name: "ok"});
okButton.enabled = false;
myWindow.add ("button", undefined, "Cancelar", {name: "cancel"});




selectButton.onClick = function () {
    if(files == null){
        files = File.openDialog("Escolha os arquivos", "*.doc;*.docx", true);
    } else{
        var temp = File.openDialog("Escolha os arquivos", "*.doc;*.docx", true);
        for(var i = 0; i < temp.length; i++){
            files.push(temp[i]);
        }
    }
    updateListBox(files,fileListBox);
}

moveUpButton.onClick = function () {
    var index = fileListBox.selection.index;
    if(moveUp(files, fileListBox.selection.index)){
        updateListBox(files,fileListBox);
        fileListBox.selection = index - 1;
    }
}

moveDownButton.onClick = function () {
    var index = fileListBox.selection.index;
    if(moveDown(files, fileListBox.selection.index)){
        updateListBox(files,fileListBox);
        fileListBox.selection = index + 1;
    }   
}

delButton.onClick = function () {
    
    var index = fileListBox.selection.index;
    
    if(fileListBox.selection != null){
        files.splice(index,1);
    }
    updateListBox(files,fileListBox);

    if(index == files.length){
        fileListBox.selection = files.length - 1;
    } else{
        fileListBox.selection = index;
    }
}

if(myWindow.show() == 1){
    if(files != null){
        if(files.length > 0){
            //currentDoc.pages.add(LocationOptions.AFTER, pages[myDropdown.selection.index])
            importFiles(files, myDropdown.selection.index);
            //linkFrames(firstPageNum)
        }
    }
}



function moveUp(arr, index){
    if(index > 0){
        var temp = arr[index-1];
        arr[index-1] = arr[index];
        arr[index] = temp;
        return true;
    } 
    return false
}

function moveDown(arr, index){
    if(index < arr.length-1){
        var temp = arr[index+1];
        arr[index+1] = arr[index];
        arr[index] = temp;
        return true;
    } 
    return false
}

function updateListBox(items, list){
    list.removeAll();
    if(items.length > 0){
        for(var i = 0; i < items.length; i++){
            list.add('item',items[i].displayName);
        }
    }
    if(items.length == 0){
        okButton.enabled = false;
        moveDownButton.enabled = false;
        moveUpButton.enabled = false;
        delButton.enabled = false;  
    } else{
        okButton.enabled = true;
        moveDownButton.enabled = true;
        moveUpButton.enabled = true;
        delButton.enabled = true;  
    }
}





// Abre a caixa de dialogo que seleciona os arquivos
function getFilesList(){
    return File.openDialog("Please choose a file", "*.doc;*.docx", true);
}

// Importa os arquivos e coloca cada um em uma página
function importFiles(files, pageNum){
    var currentPage = pages[pageNum]
    for(var i = 0; i < files.length; i++){
        currentDoc.pages.add(LocationOptions.AFTER, currentPage)
        //if(pageNum + i > currentDoc.pages.length-1){
        //    currentDoc.pages.add();
        //}
        //var currentPage = currentDoc.pages.item(pageNum + i);
        currentPage = currentDoc.pages.item(pageNum + 1);
        currentPage.place(File(files[i]), [x,y], undefined, false,false)[0];
        pageNum++;

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

    function alert_scroll (title, input){
        if (input instanceof Array)
            input = input.join ("\r");
        var w = new Window ("dialog", title);
        var list = w.add ("edittext", undefined, input, {multiline: true, scrolling: true});
        list.maximumSize.height = w.maximumSize.height-100;
        list.minimumSize.width = 250;
        w.add ("button", undefined, "Close", {name: "ok"});
        w.show ();
    }