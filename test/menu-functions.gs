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
/**  
 *	
 */
function getNamedRangesActiveCell(){
	return getNamedRanges().findNamedRanges ( getActiveRange() );
}

/**  
 *	
 */
function msgBox( msg, data)
{
	Browser.msgBox(
		msg + JSON.stringify( data )
	);
}


/*
====== MENU TEAT FUNCTIONS ======
*/

/** 
*/
function findNamedRangesMenu(){
	var nranges	= getNamedRanges().findNamedRanges (getActiveRange()).getNames();
	if(nranges===false)
		msgBox('Selected range IS NOT in named ranges' );		
	else{
		msgBox('getNamedRanges of active cell', nranges.get());
		msgBox('Array of getNamedRanges of active cell', nranges.toArray());		
	}
}

/** Test of getNamedRanges.findAllSharedgetNamedRangesMenu()
*/
function findAllSharedgetNamedRangesMenu(){
	var nranges	= getNamedRanges().findAllSharedgetNamedRanges( getActiveRange() ).getNames();
	msgBox('All getNamedRanges sharing same prefix',  nranges.get() );
	msgBox('Array of All getNamedRanges sharing same prefix',  nranges.toArray() );	
}
/** 
*/
function getPrefixesMenu(){
	//msgBox('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getPrefixes() );
	msgBox('Ranges of All getNamedRanges',  getNamedRanges().getPrefixes() );	
	msgBox('Ranges of getNamedRanges for selected range',  getNamedRanges().getPrefixes(getActiveRange()) );	

}

/** 
*/
function findPrefixesMenu(){
	var nranges_prefixes = getNamedRanges().findPrefixes( getActiveRange() ).get()
	msgBox('Prefixes of getNamedRanges for selected range',  nranges_prefixes );
}


/*
====== MENU TEAT FUNCTIONS ======
*/

/** 
*/
function getA1NotationMenu(){
	msgBox('getA1Notations of getNamedRanges for selected range',  getNamedRangesActiveCell().getA1Notation().get() );
	msgBox('getA1Notations of "names" getNamedRanges',  getNamedRanges().getA1Notation().get() );	

}
/** 
*/
function getNamesMenu(){
	msgBox('Names of getNamedRanges for selected range',  getNamedRangesActiveCell().getNames().get() );
	//msgBox('Names of prefix "names" getNamedRanges',  getNamedRanges().getNames('names').get() );	
}
/** 
*/
function getRangesMenu(){
	msgBox('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getRanges().get() );
}
/** 
*/
function getValuesMenu(){
	msgBox('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getValues().get() );
	//msgBox('Ranges of "names" getNamedRanges',  getNamedRanges().getValues('names').get() ));	
}






