/** 
 */
function addTestMenu() 
{
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('NamedRanges Test')
  
  .addItem('findNamedRanges(activeRange)', 'findNamedRangesMenu')
  .addItem('findAllSharedNamedRanges()', 'findAllSharedNamedRangesMenu')
  .addSeparator()
  .addItem('getPrefixes()  - Get all prefixes', 'getPrefixesMenu')
  .addItem('findPrefixes() - Get all prefixes for active range', 'findPrefixesMenu')
  .addSeparator()
  .addItem('getA1Notation() - For active range', 'getA1NotationMenu')
  .addItem('getNames()      - For active range', 'getNamesMenu')
  .addItem('getRanges()     - For active range', 'getRangesMenu')
  .addItem('getValues()     - For active range', 'getValuesMenu')
  
  .addToUi(); 
}
          



