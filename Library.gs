/**
 * GetTableByName()
 *
 * Read a table of data into a 2-dimensional array and optionally confirm numeric results
 */
function GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose)
{
  var spreadsheet= null;
  var range= null;
  var data= null;
  var table= [];
  var good= true;
  var maxIterations= 10;
  var sleepInterval= 5000;
  var iterationErrors= [];
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (range= spreadsheet.getRangeByName(sourceName))
    {
      for (var iteration= 0; iteration < maxIterations; iteration++)
      {
        iterationErrors= [];
        if (data= range.getValues())
        {
          for (var vIndex= 0; vIndex < data.length; vIndex++)
          {
            if (confirmNumbers)
            {
              for (var hIndex= firstDataColumn; hIndex < data[vIndex].length; hIndex++)
              {
                // check each value obtained to make sure it is a number and above the sanity check limit
                if ((data[vIndex][hIndex] == null) || isNaN(data[vIndex][hIndex]) || (data[vIndex][hIndex] < limit))
                {
                  iterationErrors.push("Could not get a viable value (<".concat(data[vIndex][hIndex], "> v. limit of <", limit,
                                                                                ") from location <", hIndex.toFixed(0), ",",  vIndex.toFixed(0),
                                                                                "> of range named <", sourceName, "> in spreadsheet <", spreadsheet.getName(), ">."));
                  
                  data[vIndex][hIndex]= iteration.toFixed(0);
                  good= false;
                }
              }
            }
            
            if (good)
            {
              // all values checked out against the limit -- save current data row and append current iteration count (if asked)
              table[vIndex]= data[vIndex];
              if (storeIterationCount)
              {
                table[vIndex].push(iteration.toFixed(0)+1);
              }
            }
          }
          
          if (good)
          {
            // all values checked out against the limit -- we're done here
            break;
          }
          else
          {
            // reset the flag -- we'll try again
            good= true;
          }
        }
        else
        {
          Logger.log("[GetTableByName] Could not read data from range named <%s> in spreadsheet <%s>.", sourceName, spreadsheet.getName());
          data= iteration;
        }
        
        Utilities.sleep(sleepInterval);
      }
      
      if (iterationErrors.length > 0)
      {
        // encountered errors while reading data -- report them
        while(iterationErrors.length > 0)
        {
          // report all the accumulated errors
          Logger.log("[GetTableByName] " + iterationErrors.shift());
        }
        Logger.log("[GetTableByName] Reached <%s> iterations but still could not get all data from range named <%s> in spreadsheet <%s>.",
                   iteration.toFixed(0), sourceName, spreadsheet.getName());
        
        table= null;
      }
    }
    else
    {
      Logger.log("[GetTableByName] Could not get range named <%s> in spreadsheet <%s>.", sourceName, spreadsheet.getName());
      table= null;
    }
  }
  else
  {
    Logger.log("[GetTableByName] Could not open spreadsheet ID <%s>.", id);
    table= null;
  }
  
  return table;
};


/**
 * GetValueByName()
 *
 * Obtain a value from a labeled one-cell range
 */
function GetValueByName(id, sourceName, verbose, confirmNumbers, limit)
{
  var value= null;
  var firstDataColumn= 0;
  var storeIterationCount= false;
  
  if (confirmNumbers == undefined)
  {
    confirmNumbers= false;
    limit= 0;
    if (verbose)
    {
      Logger.log("[SaveValues] confirmNumbers set to default <%s> with limit set to <%s>.", confirmNumbers, limit);
    }
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit= 0;
        if (verbose)
        {
          Logger.log("[SaveValues] limit set to default <%s>.", limit);
        }
      }
    }
  }
  
  value= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose); 
  if (value)
  {
    // We seem to have something!
    if (value.length > 0)
    {
      // We seem to have at least one dimension!
      if (value[0].length > 0)
      {
        // We seem to have a proper table!
        // Return the top-left value
        return value[0][0];
      }
      else
      {
        // Not a proper table!
        Logger.log("[GetValueByName] Range named <%s> is not a table.", sourceName);
        return null;
      }
    }
    else
    {
      // Not even a proper array!
      Logger.log("[GetValueByName] Range named <%s> is not even an array.", sourceName);
      return null;
    }
  }
  else
  {
    // We got nothing!
    Logger.log("[GetValueByName] Range named <%s> did not result in a viable value.", sourceName);
    return null;
  } 
};


/**
 * SetTableByName()
 *
 * Write a 2-dimensional array of data into a named spreadsheet table
 */
function SetTableByName(id, destinationName, table, verbose)
{
  var spreadsheet= null;
  var range= null;
  var success= true;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (range= spreadsheet.getRangeByName(destinationName))
    {
      // write out the values
      range= range.setValues(table);
      if (!range)
      {
        Logger.log("[SetTableByName] Could not write out range named <%s> in spreadsheet <%s>.", destinationName, spreadsheet.getName());
        success= false;
      }
    }
    else
    {
      Logger.log("[SetTableByName] Could not get range named <%s> in spreadsheet <%s>.", destinationName, spreadsheet.getName());
      success= false;
    }
  }
  else
  {
    Logger.log("[SetTableByName] Could not open spreadsheet ID <%s>.", id);
    success= false;
  }
  
  return success;
};


/**
 * SetValueByName()
 *
 * Set a value from to labeled one-cell range
 */
function SetValueByName(id, destinationName, value, verbose)
{
  return SetTableByName(id, destinationName, [[value]], verbose)
};


/**
 * GetCurrentAndPriorYearSheets()
 *
 * Looks up source and destination sheet ID
 */
function GetCurrentAndPriorYearSheets(id, currentYear, priorYear, verbose)
{
  var sourceName= "ExternalLookups";
  var firstDataColumn= 0;
  var idsByYear= [];
  var sheetIDs= {};
  var confirmNumbers= false;
  var storeIterationCount= false;
  
  idsByYear= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, storeIterationCount, verbose);
  
  if (idsByYear)
  {
    // we have viable IDs
    for (var vIndex= 0; vIndex < idsByYear.length; vIndex++)
    {
      if (idsByYear[vIndex][firstDataColumn] == currentYear)
      {
        sheetIDs["currentYear"]= idsByYear[vIndex][firstDataColumn+1];
      }
      else if (idsByYear[vIndex][firstDataColumn] == priorYear)
      {
        sheetIDs["priorYear"]= idsByYear[vIndex][firstDataColumn+1];
      }
    }
  }
  
  return sheetIDs;
};


/**
 * SaveValues()
 *
 * Save current values in a mirror table 
 */
function SaveValues(id, sourceName, destinationName, verbose, confirmNumbers, limit)
{
  var sourceValues= [];
  var destinationValues= [];
  var firstDataColumn= 0;
  var storeIterationCount= false;
  var changed= false;
  
  // set defaults unless supplied
  if (confirmNumbers == undefined)
  {
    confirmNumbers= false;
    limit= 0;
    if (verbose)
    {
      Logger.log("[SaveValues] confirmNumbers set to default <%s> with limit set to <%s>.", confirmNumbers, limit);
    }
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit= 0;
        if (verbose)
        {
          Logger.log("[SaveValues] limit set to default <%s>.", limit);
        }
      }
    }
  }
  
  // read all the source and destination values, compare, and update
  if (sourceValues= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
  {
    // we have source values, proceed to destination values
    if (destinationValues= GetTableByName(id, destinationName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
    {
      // compare values and update them
      if (sourceValues.length == destinationValues.length)
      {
        for (var vIndex= 0; vIndex < destinationValues.length; vIndex++)
        {
          if (sourceValues[vIndex].length == destinationValues[vIndex].length)
          {
            for (var hIndex= 0; hIndex < destinationValues[vIndex].length; hIndex++)
            {
              if (sourceValues[vIndex][hIndex] != destinationValues[vIndex][hIndex])
              {
                if (verbose)
                {
                  Logger.log("[SaveValues] Value at location <%s, %s> has changed to <%s> in table <%s> from <%s> in table <%s> of spreadsheet ID <%s>.",
                             hIndex.toFixed(0), vIndex.toFixed(0), sourceValues[vIndex][hIndex], sourceName, destinationValues[vIndex][hIndex], destinationName, id);
                }
                destinationValues[vIndex][hIndex]= sourceValues[vIndex][hIndex];
                changed= true;
              }
              else
              {
                if (verbose)
                {
                  Logger.log("[SaveValues] Value <%s> (<%s>) at location <%s, %s> has not changed between named tables <%s> and <%s> of spreadsheet ID <%s>.",
                             destinationValues[vIndex][hIndex], sourceValues[vIndex][hIndex], hIndex.toFixed(0), vIndex.toFixed(0), sourceName, destinationName, id);
                }
              }
            }
          }
          else
          {
            Logger.log("[SaveValues] Source values range <%s, %s> of source <%s> does not match destination range <%s, %s> of destination <%s> in sheet ID <%s>.",
                   sourceValues.length, sourceValues[vIndex].length, sourceName, destinationValues.length, destinationValues[vIndex].length, destinationName, id);
          }
        }
      }
      else
      {
        Logger.log("[SaveValues] Source values height <%s> of source <%s> does not match destination range <%s> of destination <%s> in sheet ID <%s>.",
                   sourceValues.length, sourceName, destinationValues.length, destinationName, id);
      }
      
      if (changed)
      {
        // write out the values
        if (!SetTableByName(id, destinationName, destinationValues, verbose))
        {
          Logger.log("[SaveValues] Could not write out range named <%s> in spreadsheet ID <%s>.", destinationName, id);
        }
      }
    }
    else
    {
      Logger.log("[SaveValues] Could not get range named <%s> in spreadsheet ID <%s>.", destinationName, id);
    }
  }
  else
  {
    Logger.log("[SaveValues] Could not get range named <%s> in spreadsheet ID <%s>.", sourceName, id);
  }
};


/**
 * GetLastSnapshotStamp()
 *
 * Obtain the identifier stamp for the last snapshot entry 
 */
function GetLastSnapshotStamp(id, sheetName, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      if (height= sheet.getLastRow())
      {
        if (range= sheet.getRange(height, 1))
        {
          return range.getValue();
        }
        else
        {
          Logger.log("[GetLastSnapshotStamp] Could not set range to the first cell of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                     height.toFixed(0), sheetName, spreadsheet.getName());
        }
      }
      else
      {
        Logger.log("[GetLastSnapshotStamp] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[GetLastSnapshotStamp] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[GetLastSnapshotStamp] Could not open spreadsheet ID <%s>.", id);
  }
};


/**
 * GetLastSnapshotCell()
 *
 * Obtain the last cell of the last snapshot entry 
 */
function GetLastSnapshotCell(id, sheetName, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var row= null;
  var column= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      row= sheet.getLastRow();
      column= sheet.getLastColumn();
      if (row && column)
      {
        if (range= sheet.getRange(row, column))
        {
          return range.getValue();
        }
        else
        {
          Logger.log("[GetLastSnapshotCell] Could not set range to the last cell of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                     height.toFixed(0), sheetName, spreadsheet.getName());
        }
      }
      else
      {
        Logger.log("[GetLastSnapshotCell] Could not learn the last row and column in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[GetLastSnapshotCell] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[GetLastSnapshotCell] Could not open spreadsheet ID <%s>.", id);
  }
};


/**
 * CheckSnapshot()
 *
 * Check data in the destination spreadsheet 
 */
function CheckSnapshot(id, sheetName, newDataDate, verbose)
{
  var lastDataDate= new Date(GetLastSnapshotStamp(id, sheetName, verbose));
  
  if (lastDataDate && (lastDataDate.getFullYear() == newDataDate.getFullYear()) && (lastDataDate.getMonth() == newDataDate.getMonth())
  && (lastDataDate.getDate() == newDataDate.getDate()))
  {
    // we have already recorded data for today
    return true;
  }
  else if (lastDataDate > newDataDate)
  {
    Logger.log("[CheckSnapshot] We seem to have stale data from the past (last date <%s> is later than new date <%s>), skipping...",
               lastDataDate, newDataDate);
    
    return true;
  }
  else
  {
    // we don't have the latest data
    return false; 
  }
};


/**
 * CompileSnapshot()
 *
 * Compile our snapshot from various cells
 */
function CompileSnapshot(id, names, limits, now, verbose)
{
  var snapshot= [now];
  var table= [];
  var firstDataColumn= 0;
  var confirmNumbers= true;
  var storeIterationCount= true;
  var good= true;
  var iterations= 0;
  
  for (var counter= 0; counter < names.length; counter++)
  {
    // read each table or cell of interest and accumulate in an array
    table= GetTableByName(id, names[counter], firstDataColumn, confirmNumbers, limits[counter], storeIterationCount, verbose);
    if (table)
    {
      // we got viable data, now transpose the table
      for (var row= 0; row < table.length; row++)
      {
        // grab the first column value from every row returned
        snapshot.push(table[row][0]);
        if (iterations < table[row][table[row].length-1])
        {
          // store the highest iteration count
         iterations= table[row][table[row].length-1];
        }
        
        //Logger.log("[CompileSnapshot] Adding value <%s> from row <%s> of source <%s>.", table[row][0], row.toFixed(0), names[counter]);
      }
    }
    else
    {
      // we failed to obtain viable data
      good= false;
      break;
    }
  }
  
  
  if (good)
  {
    // dress and return viable data
    snapshot[0]= new Date();
    snapshot.push(iterations);
    return snapshot;
  }
  else
  {
    // no viable data to return
    return null; 
  }
};


/**
 * SaveSnapshot()
 *
 * Save values snapshot in a history table 
 */
function SaveSnapshot(id, sheetName, values, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var success= false;
  
  if (values)
  {
    // we have viable data to save
    if (!Array.isArray(values[0]))
    {
      // we seem to have a one-dimensional array -- convert it
      values= [values];
    }
    
    // now access the spreadsheet and save
    if (spreadsheet= SpreadsheetApp.openById(id))
    {
      if (sheet= spreadsheet.getSheetByName(sheetName))
      {
        if (height= sheet.getLastRow())
        {
          if (range= sheet.getRange(height+1, 1, values.length, values[0].length))
          {
            if (range= range.setValues(values))
            {
              PropagateFormulas(sheet, height+1, values[0].length, verbose);
              success= true;
            }
            else
            {
              Logger.log("[SaveSnapshot] Could not append values in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName()); 
            }
          }
          else
          {
            Logger.log("[SaveSnapshot] Could not set range to append beyond the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                       height.toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else
        {
          Logger.log("[SaveSnapshot] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
      }
      else
      {
        Logger.log("[SaveSnapshot] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[SaveSnapshot] Could not open spreadsheet ID <%s>.", id);
    }
  }
  else
  {
    Logger.log("[SaveSnapshot] Nothing to write to sheet <%s> of spreadsheet ID <%s>...", sheetName, id);
  }
  
  return success;
};


/**
 * PropagateFormulas()
 *
 * Propagate formulas from the row above 
 */
function PropagateFormulas(sheet, row, column, verbose)
{
  var width= null;
  var formulas= null;
  var range= null;
  
  if (width= sheet.getLastColumn())
  {
    if (width > column)
    {
      // looks like we have spare columns to check
      if (formulas= sheet.getRange(row-1, column+1, 1, width-column).getFormulas())
      {
        if (range= sheet.getRange(row, column+1, 1, width-column).setFormulas(formulas))
        {
          if (verbose)
          {
            Logger.log("[PropagateFormulas] Updated formulas in columns <%s> through <%s> of row <%s> in sheet <%s>.",
                       (column+1).toFixed(0), width.toFixed(0), row.toFixed(0), sheet.getName());
          }
        }
        else
        {
          Logger.log("[PropagateFormulas] Could not set formulas in columns <%s> through <%s> of row <%s> in sheet <%s>.",
                     (column+1).toFixed(0), width.toFixed(0), row.toFixed(0), sheet.getName());
        }
      }
      else
      {
        Logger.log("[PropagateFormulas] Could not read formulas from columns <%s> through <%s> of row <%s> in sheet <%s>.",
                   (column+1).toFixed(0), width.toFixed(0), (row-1).toFixed(0), sheet.getName());
      }
    }
    else
    {
      if (verbose)
      {
        Logger.log("[PropagateFormulas] No columns to propagate in sheet <%s>.", sheet.getName());
      }
    }
  }
  else
  {
    Logger.log("[PropagateFormulas] Could not obtain width of sheet <%s>.", sheet.getName());
  }
};


/**
 * UpdateSnapshotCell()
 *
 * Update a specific value in a history table 
 */
function UpdateSnapshotCell(id, sheetName, column, value, onlyIfBlank, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var success= false;
  
  if (value != null)
  {
    // now access the spreadsheet and save
    if (spreadsheet= SpreadsheetApp.openById(id))
    {
      if (sheet= spreadsheet.getSheetByName(sheetName))
      {
        if (height= sheet.getLastRow())
        {
          if (range= sheet.getRange(height, column, 1, 1))
          {
            if (!onlyIfBlank || range.isBlank())
            {
              if (range= range.setValue([[value]]))
              {
                if (verbose)
                {
                  Logger.log("[UpdateSnapshotCell] Updated cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s> with <%s>.",
                             column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName(), value);
                }
                
                success= true;
              }
              else
              {
                Logger.log("[UpdateSnapshotCell] Could not update cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                           column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName());
              }
            }
            else
            {
              Logger.log("[UpdateSnapshotCell] Could not update cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s> "
                         + "since that would clobber an existing value <%s>.", column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName(), range.getValue());
              
            }
          }
          else
          {
            Logger.log("[UpdateSnapshotCell] Could not set range to update the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                       height.toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else
        {
          Logger.log("[UpdateSnapshotCell] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
      }
      else
      {
        Logger.log("[UpdateSnapshotCell] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[UpdateSnapshotCell] Could not open spreadsheet ID <%s>.", id);
    }
  }
  else
  {
    Logger.log("[UpdateSnapshotCell] Nothing to update in column <%s> of sheet <%s> in spreadsheet ID <%s>...", column, sheetName, id);
  }
  
  return success;
};


/**
 * SaveValuesInHistory()
 *
 * Saves current values in a history table
 */
function SaveValuesInHistory(id, sheetName, sourceNames, sourceLimits, now, backupRun, verbose)
{
  if (CheckSnapshot(id, sheetName, now, verbose))
  {
    if (!backupRun)
    {
      Logger.log("[SaveValuesInHistory] Redundant primary run at <%s> for sheet <%s>", now, sheetName);
    }
  }
  else
  {
    SaveSnapshot(id, sheetName, CompileSnapshot(id, sourceNames, sourceLimits, now, verbose), verbose);
    if (backupRun)
    {
      Logger.log("[SaveValuesInHistory] Primary run seems to have failed for sheet <%s>...", sheetName);
    }
  }
};


/**
 * RemoveDuplicateSnapshot()
 *
 * Remove a recent duplicate entry from the history table
 */
function RemoveDuplicateSnapshot(id, sheetName, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var width= null;
  var rowData= null;
  var ultimateStamp= null;
  var penultimateStamp= null;
  var priorStamp= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      if (height= sheet.getLastRow())
      {
        // Learn the latest time stamp
        if (range= sheet.getRange(height, 1))
        {
          ultimateStamp= new Date(range.getValue());
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                     height.toFixed(0), sheetName, spreadsheet.getName());
        }
        
        
        // Learn the prior time stamp (seemingly real and accurate data two rows above latest)
        if (range= sheet.getRange(height-2, 1))
        {
          priorStamp= new Date(range.getValue());
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the prior row <%s> in sheet <%s> for spreadsheet <%s>.",
                     (height-2).toFixed(0), sheetName, spreadsheet.getName());
        }
        
        // Learn the time stamp just before the latest
        if (width= sheet.getLastColumn())
        {
          if (range= sheet.getRange(height-1, 1, 1, width))
          {
            rowData= range.getValues();
            penultimateStamp= new Date(rowData[0][0]);
          }
          else
          {
            Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the next to the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                       (height-1).toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not learn the last column in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
        
        // Remove the next to the last row if its time stamp matches that of the last row or the prior row preceding it
        if (ultimateStamp.getTime() === penultimateStamp.getTime() || priorStamp.getTime() === penultimateStamp.getTime())
        {
          try
          {
            sheet= sheet.deleteRow(height-1)
          }
          catch (error)
          {
            Logger.log("[RemoveDuplicateSnapshot] Failed to remove duplicate row:\n".concat(error));
          }
          
          if (sheet)
          {
            return [[priorStamp, "Kept"], rowData[0], [ultimateStamp, "Kept"]];
          }
          else
          {
            Logger.log("[RemoveDuplicateSnapshot] Failed to remove the penultimate row for time stamp <%s> in sheet <%s> for spreadsheet <%s>.",
                       penultimateStamp, sheetName, spreadsheet.getName());
          }
        }
        else
        {
          if (verbose)
          {
            Logger.log("[RemoveDuplicateSnapshot] No need to remove history rows as time stamps (<%s> and <%s>) do not match in sheet <%s> for spreadsheet <%s>.",
                       ultimateStamp, penultimateStamp, sheetName, spreadsheet.getName());
          }
        }
      }
      else
      {
        Logger.log("[RemoveDuplicateSnapshot] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[RemoveDuplicateSnapshot] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[RemoveDuplicateSnapshot] Could not open spreadsheet ID <%s>.", id);
  }
  
  return false;
};


/**
 * TrimHistory()
 *
 * Remove earliest entries from the history table
 */
function TrimHistory(id, sheetName, maxRows, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var height= 0;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      height= sheet.getLastRow();
      
      // Accommodate the header row!
      if (height > (maxRows + 1))
      {
        try
        {
          sheet.deleteRows(2, height - (maxRows + 1));
        }
        catch (error)
        {
          Logger.log("[TrimHistory] Failed to trim history rows:\n".concat(error));
        }
      }
    }
    else
    {
      Logger.log("[TrimHistory] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[TrimHistory] Could not open spreadsheet ID <%s>.", id);
  }
};


/**
 * Synchronize()
 *
 * Preserve value from one range in another range (designed to work on first elements only)
 */
function Synchronize(sourceID, destinationID, sourceNames, destinationNames, verbose, verboseChanges)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var value= null;
  var format= null;
  var sourceValues= [];
  
  if (verboseChanges == null)
  {
    // initialize omitted optional verbose flag to match overall verbose flag
    verboseChanges= verbose;
  }
  
  if (spreadsheet= SpreadsheetApp.openById(sourceID))
  {
    for (var counter= 0; counter < sourceNames.length; counter++)
    {
      // read all the source values
      if (range= spreadsheet.getRangeByName(sourceNames[counter]))
      {
        value= range.getValue();
        format= range.getNumberFormats()[0][0];
        if (isNaN(value))
        {
          Logger.log("[Synchronize] Failed to obtain real value <%s> from range named <%s> in spreadsheet <%s>.",
                     value, sourceNames[counter], spreadsheet.getName());
          sourceValues.push(null);
        }
        else
        {
          if (format.indexOf("%") > -1)
          {
            // counteract automatic % adjustments
            //sourceValues.push((value*100).toFixed(2));
            sourceValues.push(value);
          }
          else
          {
            sourceValues.push(value.toFixed(2));
          }
        }
      }
      else
      {
        Logger.log("[Synchronize] Could not get range named <%s> in spreadsheet <%s>.", sourceNames[counter], spreadsheet.getName());
      }
    }
  }
  else
  {
    Logger.log("[Synchronize] Could not open spreadsheet ID <%s>.", sourceID);
  }
  
  if (sourceValues.length == destinationNames.length)
  {
    if (spreadsheet= SpreadsheetApp.openById(destinationID))
    {
      for (var counter= 0; counter < destinationNames.length; counter++)
      {
        // compare and write values
        if (range= spreadsheet.getRangeByName(destinationNames[counter]))
        {
          value= range.getValue();
          format= range.getNumberFormats()[0][0];
          //if (format.indexOf("%") > -1)
          //{
          //  // counteract automatic % adjustments
          //  value= (value*100).toFixed(2);
          //}
          if ((sourceValues[counter] != null) && (value != sourceValues[counter]))
          {
            // looks like the value has changed -- update it 
            range.setValue(sourceValues[counter]);
            if (verboseChanges)
            {
              Logger.log("[Synchronize] Value for range <%s> in sheet <%s> updated to <%s>, it was <%s>.",
                         destinationNames[counter], spreadsheet.getName(), sourceValues[counter], value);
            }
          }
          else
          {
            if (verbose)
            {
              Logger.log("[Synchronize] Value for range <%s> has not changed <%s>.", destinationNames[counter], sourceValues[counter]);
            }
          }
        }
        else
        {
          Logger.log("[Synchronize] Could not get range named <%s> in spreadsheet <%s>.", destinationNames[counter], spreadsheet.getName());
        }
      }
    }
    else
    {
      Logger.log("[Synchronize] Could not open spreadsheet ID <%s>.", destinationID);
    }
  }
  else
  {
    Logger.log("[Synchronize] Source values range <%s> does not match destination range <%s>.", sourceValues.length.toFixed(0), destinationNames.length.toFixed(0));
  }
};


/**
 * GetParameters()
 *
 * Read specified table and return an associative array comprised of key-value pairs from the first two columns
 */
function GetParameters(id, sourceName, verbose)
{ 
  var parameters= {"id": id, "verbose": verbose};
  var firstDataColumn= 1;
  var table= GetTableByName(id, sourceName, firstDataColumn, verbose);
  
  if (table)
  {
    // We seem to have something!
    if (table.length > 0)
    {
      // We seem to have at least one dimension!
      if (table[0].length >= 2)
      {
        // We seem to have at least two columns
        for (var row= 0; row < table.length; row++)
        {
          // Check each row for a viable key-value pair and preserve them in our associative array 
          if (table[row][0] != null && table[row][1] != null)
          {
            // We seem to have a viable key-value pair
            parameters[table[row][0]]= table[row][1];
          }
        }
      }
      else
      {
        // Not a proper table!
        Logger.log("[GetParameters] Range named <%s> is not a table.", sourceName);
      }
    }
    else
    {
      // Not even a proper array!
      Logger.log("[GetParameters] Range named <%s> is not even an array.", sourceName);
    }
  }
  else
  {
    // We got nothing!
    Logger.log("[GetParameters] Range named <%s> did not result in a viable value.", sourceName);
  }
  
  return parameters;
};


/**
 * GetMainSheetID()
 *
 * Return a one-dimensional array willed with specified values
 */
function GetMainSheetID()
{ 
  return "18pUgp_50UsEd-bh-NF6rH5Qifu2T-XsngyXg9DvElnw";
};


/**
 * FillArray()
 *
 * Return a one-dimensional array willed with specified values
 */
function FillArray(size, value)
{
  var fill= [];
  var counter= 0;
  
  while (counter < size)
  {
    fill[counter++] = value;
  }
  
  return fill;
};


/**
 * NumberToString()
 *
 * Return a formatted number as a string
 */
function NumberToString(number, width, pad)
{
  var formattedNumber= "" + number;
  
  while (formattedNumber.length < width)
  {
    formattedNumber= pad + formattedNumber;
  }
  
  return formattedNumber;
};


/**
 * DateToLocaleString()
 *
 * Return a formatted date as a short string (m/dd/yyyy hh:mm:ss)
 */
function DateToLocaleString(date, separator)
{
  var dateOptions= { day: '2-digit', month: '2-digit', year: 'numeric' };
  //var timeOptions= { hour12: false, hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  var timeOptions= { hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  
  if (date == undefined)
  {
    date= new Date();
  }
  else
  {
    date= new Date(date);
  }
  
  if (separator == undefined)
  {
    separator= " "; 
  }
  
  //return date.toLocaleString('en-US', {hour12: false, hourCycle: 'h23'});
  //return date.toLocaleString('en-US', {hourCycle: 'h23'});
  return date.toLocaleDateString('en-US', dateOptions) + separator + date.toLocaleTimeString('en-US', timeOptions)
};


/**
 * DateToStringShort()
 *
 * Return a formatted date as a short string (yyyy-mm-dd hh:mm:ss)
 *
 * https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
 * https://developers.google.com/apps-script/reference/utilities/utilities#formatdatedate-timezone-format
 *
function DateToStringShort(id, date)
{
  return Utilities.formatDate(date, GetTimeZone(id), "yyyy-MM-dd HH:mm:ss Z");
};


/**
 * GetTimeZone()
 *
 * Return time zone associated with the spreadsheet in question
 *
function GetTimeZone(id)
{
  // declare constants and local variables
  var spreadsheet= null;
  var timeZone= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (timeZone= spreadsheet.getSpreadsheetTimeZone())
    {
      return timeZone;
    }
  }
  
  return "America/Chicago";
};

*/