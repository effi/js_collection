function createEmail(subject, body){
	try{
		var theApp = new ActiveXObject("Outlook.Application");
		var objNS = theApp.GetNameSpace('MAPI');
		var theMailItem = theApp.CreateItem(0); // value 0 = MailItem
		return theMailItem;
	}
	catch (err) {
	 alert(err.message);
	} 
}
