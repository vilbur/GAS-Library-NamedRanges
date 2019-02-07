/*=====================================================================================================================================*/
/*                                                                                                                                     */
/*  getNamedRanges - MENU FUNCTIONS                                                                                                    */
/*  Contetn of this file has to be in scope of MAIN SCRIPT                                                                             */
/*                                                                                                                                     */
/*=====================================================================================================================================*/

function getNamedRanges(){
	return new NamedRanges();
}

/* ================================================================= */
/*	HELP DIALOG                                                      */
/* ================================================================= */
/**
 */
function NamedRangesDialogHelp() {
	getNamedRanges().showDialog('NamedRanges-DialogHelp.html', project_name + ' Help');
}


/* ================================================================= */
/*	MENU FUNCTIONS                                                   */
/* ================================================================= */

/*
====== HELPER FUNCTIONS ======
*/


/** getActiveRange
*/
function getActiveRange(){
	return SpreadsheetApp.getActive().getActiveRange();
	//return SpreadsheetApp.getActive().getRange('A2');
}
function getNamedRangesActiveCell(){
	return getNamedRanges().findNamedRanges ( getActiveRange() );
}

/*
====== MENU TEAT FUNCTIONS ======
*/

/** 
*/
function findNamedRangesMenu(){
	var nranges	= getNamedRanges().findNamedRanges (getActiveRange()).getNames();
	if(nranges===false)
		Browser.msgBox('Selected range IS NOT in named ranges' );		
	else{
		Browser.msgBox('getNamedRanges of active cell: '+JSON.stringify(nranges.get()));
		Browser.msgBox('Array of getNamedRanges of active cell: '+JSON.stringify(nranges.toArray()));		
	}
}

/** Test of getNamedRanges.findAllSharedgetNamedRangesMenu()
*/
function findAllSharedgetNamedRangesMenu(){
	var nranges	= getNamedRanges().findAllSharedgetNamedRanges( getActiveRange() ).getNames();
	Browser.msgBox('All getNamedRanges sharing same prefix: '+JSON.stringify( nranges.get() ));
	Browser.msgBox('Array of All getNamedRanges sharing same prefix: '+JSON.stringify( nranges.toArray() ));	
}
/** 
*/
function getPrefixesMenu(){
	//Browser.msgBox('Ranges of getNamedRanges for selected range: '+JSON.stringify( getNamedRangesActiveCell().getPrefixes() ));
	Browser.msgBox('Ranges of All getNamedRanges: '+JSON.stringify( getNamedRanges().getPrefixes() ));	
	Browser.msgBox('Ranges of getNamedRanges for selected range: '+JSON.stringify( getNamedRanges().getPrefixes(getActiveRange()) ));	

}

/** 
*/
function findPrefixesMenu(){
	var nranges_prefixes = getNamedRanges().findPrefixes( getActiveRange() ).get()
	Browser.msgBox('Prefixes of getNamedRanges for selected range: '+JSON.stringify( nranges_prefixes ));
}


/*
====== MENU TEAT FUNCTIONS ======
*/

/** 
*/
function getA1NotationMenu(){
	Browser.msgBox('getA1Notations of getNamedRanges for selected range: '+JSON.stringify( getNamedRangesActiveCell().getA1Notation().get() ));
	Browser.msgBox('getA1Notations of "names" getNamedRanges: '+JSON.stringify( getNamedRanges().getA1Notation().get() ));	

}
/** 
*/
function getNamesMenu(){
	Browser.msgBox('Names of getNamedRanges for selected range: '+JSON.stringify( getNamedRangesActiveCell().getNames().get() ));
	//Browser.msgBox('Names of prefix "names" getNamedRanges: '+JSON.stringify( getNamedRanges().getNames('names').get() ));	
}
/** 
*/
function getRangesMenu(){
	Browser.msgBox('Ranges of getNamedRanges for selected range: '+JSON.stringify( getNamedRangesActiveCell().getRanges().get() ));
}
/** 
*/
function getValuesMenu(){
	Browser.msgBox('Ranges of getNamedRanges for selected range: '+JSON.stringify( getNamedRangesActiveCell().getValues().get() ));
	//Browser.msgBox('Ranges of "names" getNamedRanges: '+JSON.stringify( getNamedRanges().getValues('names').get() ));	
}






