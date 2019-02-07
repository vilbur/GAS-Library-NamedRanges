/**
 */
function NamedRangesDialogHelp() {
	getNamedRanges().showDialog('NamedRanges-DialogHelp.html', project_name + ' Help');
}

/*
====== HELPER FUNCTIONS ======
*/

function getNamedRanges()
{
	return new NamedRanges();
}

/** getActiveRange
*/
function getActiveRange()
{
	return SpreadsheetApp.getActive().getActiveRange();
	//return SpreadsheetApp.getActive().getRange('A2');
}
/**  
 *	
 */
function getNamedRangesActiveCell()
{
	return getNamedRanges().findNamedRanges ( getActiveRange() );
}

/**  
 *	
 */
function namedRangesTestResult( msg, data)
{
	Browser.msgBox(
		msg + JSON.stringify( data )
	);
}

/*
====== MENU TEST FUNCTIONS ======
*/

/** 
*/
function findNamedRangesMenu()
{
	var nranges	= getNamedRanges().findNamedRanges (getActiveRange()).getNames();
	if(nranges===false)
		namedRangesTestResult('Selected range IS NOT in named ranges' );		
	else{
		namedRangesTestResult('getNamedRanges of active cell', nranges.get());
		namedRangesTestResult('Array of getNamedRanges of active cell', nranges.toArray());		
	}
}

/** Test of getNamedRanges.findAllSharedgetNamedRangesMenu()
*/
function findAllSharedgetNamedRangesMenu()
{
	var nranges	= getNamedRanges().findAllSharedgetNamedRanges( getActiveRange() ).getNames();
	namedRangesTestResult('All getNamedRanges sharing same prefix',  nranges.get() );
	namedRangesTestResult('Array of All getNamedRanges sharing same prefix',  nranges.toArray() );	
}
/** 
*/
function getPrefixesMenu()
{
	//namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getPrefixes() );
	namedRangesTestResult('Ranges of All getNamedRanges',  getNamedRanges().getPrefixes() );	
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRanges().getPrefixes(getActiveRange()) );	
}

/** 
*/
function findPrefixesMenu()
{
	var nranges_prefixes = getNamedRanges().findPrefixes( getActiveRange() ).get()
	namedRangesTestResult('Prefixes of getNamedRanges for selected range',  nranges_prefixes );
}


/** 
*/
function getA1NotationMenu()
{
	namedRangesTestResult('getA1Notations of getNamedRanges for selected range',  getNamedRangesActiveCell().getA1Notation().get() );
	namedRangesTestResult('getA1Notations of "names" getNamedRanges',  getNamedRanges().getA1Notation().get() );	
}
/** 
*/
function getNamesMenu()
{
	namedRangesTestResult('Names of getNamedRanges for selected range',  getNamedRangesActiveCell().getNames().get() );
	//namedRangesTestResult('Names of prefix "names" getNamedRanges',  getNamedRanges().getNames('names').get() );	
}
/** 
*/
function getRangesMenu()
{
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getRanges().get() );
}
/** 
*/
function getValuesMenu()
{
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getValues().get() );
	//msgBox('Ranges of "names" getNamedRanges',  getNamedRanges().getValues('names').get() ));	
}