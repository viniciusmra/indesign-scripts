function showProps(reflectProperties){
    
    
    //var frame = app.documents[0].pages[0].textFrames.add ({geometricBounds: [0, 0, 40, 40], contents: "x x"});

    // object
    //var reflectProperties = frame.parentPage.name;

    // display properties
    alert_scroll("Object properties", reflectProperties.reflect.properties.sort());
    //alert(frame.parentPage.name)
    // asdas @include showProps.jsx
}
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