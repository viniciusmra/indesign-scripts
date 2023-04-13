function main() {
	var w = new Window("dialog" , "Script by LFCorullón");
	var e = w.add("edittext" , [0,0,140,24] , "");
		
	e.addEventListener("keyup" , keypress);
	function keypress(k) {
		str = [];
		if (e.altKey) str.push("Alt");
		if (e.ctrlKey) str.push("Ctrl");
		if (e.shiftKey) str.push("Shift");
		if (e.metaKey) str.push("Win");
		str.push(k.keyName);
		alert(str);
	}
	w.show();
}

    main();