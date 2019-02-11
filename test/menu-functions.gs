/**  
 */
var MENU_NRANGE_TEST = SpreadsheetApp.getUi().createMenu('NamedRanges Test');

/*
====== TEST MENU FUNCTIONS ======
*/

/** 
*/
function findNamedRanges()
{
	var nranges	= getNamedRanges().findNamedRanges (getActiveRange()).getNames();
	
	if(nranges===false)
		namedRangesTestResult('Selected range IS NOT in named ranges' );		
	else{
		namedRangesTestResult('getNamedRanges of active cell', nranges.get());
		namedRangesTestResult('Array of getNamedRanges of active cell', nranges.toArray());		
	}
}
MENU_NRANGE_TEST.addItem('findNamedRanges(activeRange)', 'findNamedRanges');

/** Test of getNamedRanges.findAllSharedgetNamedRangesMenu()
*/
function findAllSharedgetNamedRanges()
{
	var nranges	= getNamedRanges().findAllSharedgetNamedRanges( getActiveRange() ).getNames();
	
	namedRangesTestResult('All getNamedRanges sharing same prefix',  nranges.get() );
	namedRangesTestResult('Array of All getNamedRanges sharing same prefix',  nranges.toArray() );	
}
MENU_NRANGE_TEST.addItem('findAllSharedNamedRanges()', 'findAllSharedNamedRanges').addSeparator();

/** 
*/
function getPrefixes()
{
	//namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getPrefixes() );
	namedRangesTestResult('Ranges of All getNamedRanges',  getNamedRanges().getPrefixes() );	
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRanges().getPrefixes(getActiveRange()) );	
}
MENU_NRANGE_TEST.addItem('getPrefixes()  - Get all prefixes', 'getPrefixes');

/** 
*/
function findPrefixes()
{
	var nranges_prefixes = getNamedRanges().findPrefixes( getActiveRange() ).get();
	
	namedRangesTestResult('Prefixes of getNamedRanges for selected range',  nranges_prefixes );
}
MENU_NRANGE_TEST.addItem('findPrefixes() - Get all prefixes for active range', 'findPrefixes').addSeparator();


/** 
*/
function getA1Notation()
{
	namedRangesTestResult('getA1Notations of getNamedRanges for selected range',  getNamedRangesActiveCell().getA1Notation().get() );
	namedRangesTestResult('getA1Notations of "names" getNamedRanges',  getNamedRanges().getA1Notation().get() );	
}
MENU_NRANGE_TEST.addItem('getA1Notation() - For active range', 'getA1Notation');

/** 
*/  
function getNames()
{
	namedRangesTestResult('Names of getNamedRanges for selected range',  getNamedRangesActiveCell().getNames().get() );
	//namedRangesTestResult('Names of prefix "names" getNamedRanges',  getNamedRanges().getNames('names').get() );	
}
MENU_NRANGE_TEST.addItem('getNames()      - For active range', 'getNames');

/** 
*/
function getRanges()
{
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getRanges().get() );
}
MENU_NRANGE_TEST.addItem('getRanges()     - For active range', 'getRanges');

/** 
*/
function getValues()
{
	namedRangesTestResult('Ranges of getNamedRanges for selected range',  getNamedRangesActiveCell().getValues().get() );
	//msgBox('Ranges of "names" getNamedRanges',  getNamedRanges().getValues('names').get() ));	
}
MENU_NRANGE_TEST.addItem('getValues()     - For active range', 'getValues');

/** 
 */
function addNrangeTestMenu() 
{
	MENU_NRANGE_TEST.addToUi(); 
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





