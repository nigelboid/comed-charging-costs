/**
 * Main entry point for the continuous check
 *
 * Checks for fresh data and saves it once it is available 
 */
function RunComEdFrequently()
{
  // Declare constants and local variables
  var verbose= false;
  var parameters= GetParametersComed(verbose);
  var priceMovingAverage= null;
  var priceTrend= null;
  var lastStamp= GetLatestTimeStamp(parameters);
  var clearSemaphore= false;
  var success= false;
  
  // Since ComEd provides 5-minutes prices and provides them about 8 minutes late,
  //   only run if more than 8 minutes have elapsed since the last entry and if not already running
  if (lastStamp && parameters["scriptTime"] > (lastStamp + parameters["checkThreshold"]))
  {
    UpdateStatus(parameters, "Getting prices...");
    var prices= GetPricesComed(parameters, lastStamp);
    
    //Logger.log("[RunComEdFrequently] Stale prices stamp (%s)", parameters["noNewPricesAlert"].toFixed(0));
    
    if (prices.length > 0)
    {
      // Looks like we have data, proceed
      UpdateStatus(parameters, "Getting prices: Latest prices obtained!");
      
      if (success= PrepareToCommit(parameters))
      {
        if (success= SavePrices(parameters, prices))
        {
          if (parameters["noNewPricesAlert"])
          {
            Logger.log("[RunComEdFrequently] New prices updated after a delay (%s minutes)", ConvertMillisecondsToMinutes(parameters["scriptTime"] - lastStamp));
            ClearMissingPricesAlertStamp(parameters);
          }
            
          if (success= UpdateComputedValues(parameters))
          {
            success= Notify(parameters);
          }
        }
      }  
      
      // Clean up
      if (success)
      {
        // Complete and successful new price commitment: attempt to clear semaphore
        success= ClearSemaphore(parameters);
        if(!success)
        {
          Logger.log("[RunComEdFrequently] Failed to clear current semaphore (%s) upon completion!", parameters["scriptTime"].toFixed(0));
        }
      }
      else
      {
        // Report early and abnormal termination
        success= Scuttle(parameters, "RunComEdFrequently");
      }
    }
    else
    {
      var action= "No new prices obtained";
      
      if (parameters["scriptTime"] > (lastStamp + parameters["checkThreshold"] * 2))
      {
        if (parameters["noNewPricesAlert"].length == 0)
        {
          Logger.log("[RunComEdFrequently] No new prices available (%s minutes)", ConvertMillisecondsToMinutes(parameters["scriptTime"] - lastStamp));
          UpdateMissingPricesAlertStamp(parameters);
        }
        
        UpdateStatus(parameters, action + ": Stale prices!");
      }
      else
      {
        UpdateStatus(parameters, action + ": Scrubbing history while waiting for updated prices...");
        ScrubHistory(parameters, action);
      }
    }
  }
  else
  {
    var action= "Invoked too soon";
    
    UpdateStatus(parameters, action + ": Scrubbing history instead of checking prices...");
    ScrubHistory(parameters, action);
  }
  
  LogSend(parameters["id"]);
  return success;
};


/**
 * GetParametersComed()
 *
 * Create and return a hash of various parameters 
 */
function GetParametersComed(verbose)
{
  var id= "154A_fDu9CG4O7T2H3M5dGHJfXFqqAE3OTZ5c0IXo57g";
  var parameters= GetParameters(id, "comedParameters", verbose);

  parameters["scriptTime"]= new Date().getTime();
  
  parameters["confirmNumbers"]= true;
  parameters["confirmNumbersLimit"]= -1000000;
  
  parameters["indexTime"]= 0;
  parameters["indexPrice"]= parameters["indexTime"] + 1;
  parameters["indexMovingAverage"]= parameters["indexPrice"] + 1;
  parameters["indexTrend"]= parameters["indexMovingAverage"] + 1;
  parameters["indexStamp"]= parameters["indexTrend"] + 1;
  parameters["indexStampTime"]= parameters["indexStamp"] + 1;
  parameters["indexAlert"]= parameters["indexStampTime"] + 1;
  parameters["priceTableWidth"]= parameters["indexAlert"] + 1;
  
  if (verbose)
  {
    Logger.log("[GetParametersComed] Parameters:\n\n%s", parameters);
  }
  
  return parameters;
};



/**
 * GetPricesComed()
 *
 * Grab pricing data from ComEd for the specified time interval 
 */
function GetPricesComed(parameters, intervalStart)
{
  var timeKey= parameters["comedKeyTime"];
  var priceKey= parameters["comedKeyPrice"];
  var urlHead= parameters["comedURLHead"];
  var urlDateStart= parameters["comedURLDateRangeStart"];
  var urlDateEnd= parameters["comedURLDateRangeEnd"];
  var intervalEnd= new Date(parameters["scriptTime"]);
  
  // Formulate starting and ending date stamps, offset latest stamp by one minute
  intervalStart= new Date(intervalStart + 60000);
  
  urlDateStart+= NumberToString(intervalStart.getFullYear(), 4, "0") + NumberToString(intervalStart.getMonth()+1, 2, "0")
  + NumberToString(intervalStart.getDate(), 2, "0") + NumberToString(intervalStart.getHours(), 2, "0") + NumberToString(intervalStart.getMinutes(), 2, "0");
  
  urlDateEnd+= NumberToString(intervalEnd.getFullYear(), 4, "0") + NumberToString(intervalEnd.getMonth()+1, 2, "0")
  + NumberToString(intervalEnd.getDate(), 2, "0") + NumberToString(intervalEnd.getHours(), 2, "0") + NumberToString(intervalEnd.getMinutes(), 2, "0");
  
  
  // Obtain and parse missing pricing data
  var options= {'muteHttpExceptions' : true };
  var responseOk= 200;
  var url= urlHead + urlDateStart + urlDateEnd;
  
  UpdateURL(parameters, url);
  
  var response= UrlFetchApp.fetch(url, options);
  var responseCode= response.getResponseCode();
  var priceTable= [];
  
  if (responseCode == responseOk)
  {
    // looks like we received a benign response
    var prices= JSON.parse(response.getContentText());
    var row= null;
    var timeStamp= null;
    
    prices.sort(function(a, b){return a[timeKey] - b[timeKey]});
    for (var entry in prices)
    {
      // convert each data line into a padded table row
      row= FillArray(parameters["priceTableWidth"], "");
      timeStamp= new Date();
      timeStamp.setTime(prices[entry][timeKey]);
      row[parameters["indexTime"]]= timeStamp;
      row[parameters["indexPrice"]]= prices[entry][priceKey];
      row= priceTable.push(row);
    }
    
    if (row > 0)
    {
      // Update status with latest price and time
      UpdateLastPrice(parameters, prices[row-1][priceKey]); 
    }
  }
  else
  {
    // looks like we did not obtain our prices
    if (parameters["verbose"])
    {
      Logger.log("[GetPricesComed] Asked for latest prices, but received an unexpected response code instead: <%s>", responseCode.toFixed(0));
    }
  }
  
  return priceTable;
};


/**
 * SendPriceAlert()
 *
 * Send a price alert, depending on conditions 
 */
function SendPriceAlert(parameters)
{
  var priceCurrent= parameters["priceLast"];
  var priceMovingAverage= parameters["priceMovingAverage"];
  var priceTrend= parameters["priceRegressionSlope"];;
  
  var priceExpensive= parameters["priceLimitExpensive"];
  var priceNormalUpper= parameters["priceLimitNormalUpper"];
  var priceNormalLower= parameters["priceLimitNormalLower"];
  var priceCheap= parameters["priceLimitCheap"];
  var priceThresholdMovingAverage= parameters["priceThresholdMovingAverage"];
  var priceThresholdStable= parameters["priceThresholdStable"];
  var priceAlertLast= parameters["priceAlertLast"];
  var priceThresholdAlert= parameters["priceThresholdAlert"];
  var priceAlertDNDStart= parameters["priceAlertDNDStart"];
  var priceAlertDNDEnd= parameters["priceAlertDNDEnd"];
 
  var alert= priceCurrent + "¢ (Δ= " + (priceCurrent-priceAlertLast).toFixed(1) + "¢ [±" + priceThresholdAlert + "¢]) and ";
  var status= "No one should ever see this!";
  
  UpdateStatus(parameters, "Composing trend missive...");
  
  if ((priceCurrent - priceMovingAverage) > priceThresholdStable)
  {
    alert+= "rising";
  }
  else if ((priceCurrent - priceMovingAverage) < -priceThresholdStable)
  {
    alert+= "falling";
  }
  else
  {
    alert+= "steady";
  }
  alert+= " (inflection= " + (priceTrend * parameters["priceRegressionSlope"]).toFixed(0);
  alert+= ", MA deviation= " + (priceCurrent - priceMovingAverage).toFixed(2) + "¢ [±" + priceThresholdMovingAverage + "¢]" + ")";
  
  if ((Math.abs(priceCurrent-priceAlertLast) > priceThresholdAlert)
      && ((priceTrend * parameters["priceRegressionSlope"] <= 0) || (Math.abs(priceCurrent - priceMovingAverage) > priceThresholdMovingAverage)))
  {
    // Price jumping too much since last alert and either inflecting trend (slope) or deviating too far from the moving average -- check if that merits an alert...
    // ...may trigger on a flat slope -- that's a feature!
    UpdateStatus(parameters, "Composing alert...");
    
    if (priceCurrent > priceExpensive)
    {
      // High price range
      alert= "Expensive electricity: " + alert;
    }
    else if (priceCurrent < priceCheap)
    {
      // Low price range
      alert= "Cheap electricity: " + alert;
    }
    else
    {
      // Situation within or near normal bounds -- suppress alert???
      if (priceCurrent < priceNormalLower)
      {
        // Just below normal
        alert= "Now just below normal: " + alert;
      }
      else if (priceCurrent > priceNormalUpper)
      {
        // Just above normal
        alert= "Now just above normal: " + alert;
      }
      else
      {
        // Normal
        alert= "Now normal: " + alert;
      }
    }
  }
  else
  {
    // Prices not moving much within current bounds -- suppress alert
    alert= "No alert triggered (parameters stable): " + alert;
    if (!parameters["verbose"])
    {
      status= alert;
      alert= null;
    }
  }
  
  if (alert)
  {
    // Alert condition reached -- apply cosmetics and check suppressed alert window
    var alertTime= new Date();
    var hour= alertTime.getHours();
    
    UpdateStatus(parameters, "Alert composed.");
    UpdateAlertStatus(parameters, priceCurrent, alert, alertTime);
    status= alert;
    
    if (priceAlertDNDStart > priceAlertDNDEnd)
    {
      // Adjust for crossing the day boundary
      priceAlertDNDEnd+= 24;
      if (hour < priceAlertDNDStart)
      {
        hour+= 24;
      }
    }
    
    if ((hour >= priceAlertDNDStart) && (hour < priceAlertDNDEnd))
    {
      // Suppress an actual alert during the Do Not Disturb Window, but annotate the status
      status+= " [suppressed]";
    }
    else
    {
      // Trigger an actual alert since we are outside the Do Not Disturb window
      Logger.log("[SendPriceAlert] %s.", alert);
      if (parameters["verbose"])
      {
        Logger.log("[SendPriceAlert] Current hour [%s] is outside the do not disturb window [%s - %s].", hour, priceAlertDNDStart, priceAlertDNDEnd);
      }
    }
  }
  
  return status;
};


/**
 * GetLatestTimeStamp()
 *
 * Get the latest time stamp from the history table 
 */
function GetLatestTimeStamp(parameters)
{
  var stamp= GetLastSnapshotStamp(parameters["id"], parameters["comedSheetPrices"], parameters["verbose"]);
  
  if (stamp && (stamp.toString().length > 0))
  {
    return new Date(stamp).getTime(); 
  }
  else
  {
    Logger.log("[GetLatestTimeStamp] Retrieved an invalid stamp [%s].", stamp);
    SetLatestTimeStamp(parameters);
  }
  
  return null;
};


/**
 * SetLatestTimeStamp()
 *
 * Set the latest time stamp in the history table (error condition recovery)
 */
function SetLatestTimeStamp(parameters)
{
  var stamp= null;
  var onlyIfBlank= false;
  
  if (stamp= new Date())
  {
    if (UpdateSnapshotCell(parameters["id"], parameters["comedSheetPrices"], parameters["indexTime"] + 1, stamp, onlyIfBlank, parameters["verbose"]))
    {
      Logger.log("[SetLatestTimeStamp] Overwrote latest time stamp with [%s].", stamp);
    }
  }
  else
  {
    Logger.log("[SetLatestTimeStamp] Could not overwrite latest time stamp.");
  }
};


/**
 * ScrubHistory()
 *
 * Remove duplicate rows from history (due to superseded runs?) and otherwise keep the history to a maximum number of entries
 */
function ScrubHistory(parameters, action)
{
  var scrubbedData= null;
  var semaphore= null;
  var maxRows= 3000;
  
  if (semaphore= GetSemaphore(parameters))
  {
    // Semaphore precludes scrubbing -- preserve its value for the chain of command and restore prior status
    var statusAction= "Deferred";
      
    PreserveStatus(parameters, statusAction);
    Logger.log("[ScrubHistory] Deferring scrubbing history (%s)...", SemaphoreConflictDetails(parameters, semaphore));
  }
  else
  {
    // Check for a duplicate snashot row and preserve its values for the chain of command
    if (scrubbedData= RemoveDuplicateSnapshot(parameters["id"], parameters["comedSheetPrices"], parameters["verbose"]))
    {
      UpdateStatus(parameters, "Removed duplicate history row.");
      Logger.log("[ScrubHistory] Removed duplicate history row\n\n%s", scrubbedData);
    }
    else
    {
      TrimHistory(parameters["id"], parameters["comedSheetPrices"], maxRows, parameters["verbose"]);
      UpdateStatus(parameters, action + ".");
    }
  }
  
  return scrubbedData;
};


/**
 * IsSupreme()
 *
 * Determine if another run has superseded this one (via run time stamps)
 */
function IsSupreme(parameters)
{
  var current= GetValueByName(parameters["id"], "statusRunCurrent", parameters["verbose"]);
  
  return (parameters["scriptTime"] >= current);   
}


/**
 * Superseded()
 *
 * Report that this run has been superseded by another
 */
function Superseded(parameters, caller, activity)
{
  var current= GetValueByName(parameters["id"], "statusRunCurrent", parameters["verbose"]);
  var statusMessage= "Superseded!";
  var logMessage= "";
  
  // Report a stale run
  logMessage= "Superseded by a later run (" + current.toFixed(0) + " started " + ConvertMillisecondsToMinutes(current - parameters["scriptTime"])
  + " minutes later; current= " + parameters["scriptTime"].toFixed(0) + ")";
  
  if (activity)
  {
    // Insert status information into status and log missives 
    statusMessage= activity + ": " + statusMessage;
    logMessage+= " while " + activity.toLowerCase() + "."; 
  }
  
  if (caller)
  {
    // Prepend caller indentifier to log missive
    logMessage= "[" + caller + "] " + logMessage;
  }
  
  UpdateStatus(parameters, statusMessage);
  Logger.log(logMessage);
};


/**
 * GetSemaphore()
 *
 * Obtain the latest semaphore, if any
 */
function GetSemaphore(parameters)
{
  return GetValueByName(parameters["id"], "semaphore", parameters["verbose"]);;
};


/**
 * SetSemaphore()
 *
 * Set a semaphore since this run is writing to history
 */
function SetSemaphore(parameters)
{
  var semaphore= null;
  
  semaphore= GetSemaphore(parameters);
  if (semaphore)
  {
    // Blocked by another run
    Logger.log("[SetSemaphore] Could not set semaphore (%s)!", SemaphoreConflictDetails(parameters, semaphore));
    Logger.log("[SetSemaphore] Prior status: %s", parameters["status"]);
    
    if (parameters["forceThreshold"] && (parameters["scriptTime"] - semaphore) > parameters["forceThreshold"])
    {
      // Enough time has elapsed for us to forcefully clear a prior semaphore
      var force= true;
      
      UpdateStatus(parameters, "Forcefully clearing prior semaphore...");
      Logger.log("[SetSemaphore] Attempting to forcefully clear prior semaphore (%s)...", semaphore.toFixed(0));
      
      force= ClearSemaphore(parameters, force);
      if (force)
      {
        var action= "Forcefully cleared prior semaphore";
        
        UpdateStatus(parameters, action + ": Scrubbing history...");
        // Stale semaphore cleared -- check for cobbled data
        ScrubHistory(parameters, action)
      }
      else
      {
        Logger.log("[SetSemaphore] Failed to forcefully clear prior semaphore!");
      }
    }
    
    return false;
  }
  else
  {
    // Clear to proceed
    SetValueByName(parameters["id"], "semaphore", parameters["scriptTime"].toFixed(0), parameters["verbose"]);
    SetValueByName(parameters["id"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
    parameters["semaphore"]= parameters["scriptTime"];
    
    return true;
  }
};


/**
 * ClearSemaphore()
 *
 * Clear our semaphore (writing)
 */
function ClearSemaphore(parameters, force)
{
  var semaphore= GetSemaphore(parameters);
  var success= false;
  
  if (semaphore)
  {
    // There is a semaphore set -- confirm it is current and proceed accordingly
    if (IsSupreme(parameters))
    {
      // Proceed to clear semaphore
      if (parameters["scriptTime"] == semaphore)
      {
        // Normally only clear own semaphore
        success= SetValueByName(parameters["id"], "semaphore", "", parameters["verbose"]);
        SetValueByName(parameters["id"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
      }
      else if (force)
      {
        // Clear another run's semaphore if set to do so
        success= SetValueByName(parameters["id"], "semaphore", "", parameters["verbose"]);
        SetValueByName(parameters["id"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
        Logger.log("[ClearSemaphore] Clearing a semaphore from another run (%s)!", semaphore.toFixed(0));
      }
      else
      {
        success= false;
        Logger.log("[ClearSemaphore] Cannot clear a semaphore from another run (%s)!", SemaphoreConflictDetails(parameters, semaphore));
      }
    }
    else
    {
      // Technically, should never enter this clause due to semaphore and precedence (superseded) checks -- or so I had thought!
      success= false;
      Logger.log("[ClearSemaphore] Cannot clear a semaphore since another run superseded this one (%s)!", SemaphoreConflictDetails(parameters, semaphore));
    }
  }
  else
  {
    // No semaphore set?! 
    success= false;
    Logger.log("[ClearSemaphore] Something or someone else has already cleared the semaphore (%s)!", parameters["scriptTime"].toFixed(0));
  }
  
  return success;
};


/**
 * Scuttle()
 *
 * Deal with an abnormal and early termination
 */
function Scuttle(parameters, caller)
{
  var success= false;
  var logMessage= "Scuttling";
  var statusAction= "Scuttled";
  
  UpdateStatus(parameters, "Scuttling...");
  
  // Compose and commit log message
  if (caller != undefined)
  {
    logMessage= "[" + caller + "] " + logMessage;
  }
  
  if (parameters["activity"])
  {
    logMessage+= " while " + parameters["activity"] + ".";
  }
  else
  {
    logMessage+= ".";
  }
  
  Logger.log(logMessage);
  
  // Clean up
  if (parameters["semaphore"] == parameters["scriptTime"])
  {
    success= ClearSemaphore(parameters);
    if (!success)
    {
      Logger.log("[Scuttle] Failed to clear current semaphore (%s)!", parameters["scriptTime"].toFixed(0));
    }
    else
    {
      // Scuttling is a failure regardless of intermediate successes!
      success= false;
    }
  }
  
  // Restore prior status for consistency and report failure
  PreserveStatus(parameters, statusAction);
  return success;
};


/**
 * SemaphoreConflictDetails()
 *
 * Compose string descricing semaphore conflict details
 */
function SemaphoreConflictDetails(parameters, semaphore)
{
  var deltaValue= ConvertMillisecondsToMinutes(parameters["scriptTime"] - semaphore);
  var details= "blocked by semaphore ";
  
  details= details.concat(semaphore.toFixed(0), " set ");
  
  if (deltaValue < 0)
  {
    // Conflicting semaphore set later
    details= details.concat(-deltaValue, " minutes later");
  }
  else
  {
    // Conflicting semaphore set earlier or concurrently
    details= details.concat(deltaValue, " minutes earlier");
  }
  
  details= details.concat("; current: ", parameters["scriptTime"].toFixed(0));
  
  return details;
};


/**
 * UpdateStatus()
 *
 * Update status and time
 */
function UpdateStatus(parameters, status)
{
  SetValueByName(parameters["id"], "status", status, parameters["verbose"]);
  SetValueByName(parameters["id"], "statusTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * PreserveStatus()
 *
 * Preserve previous status
 */
function PreserveStatus(parameters, statusAction)
{
  var statusPrior= parameters["status"];
  var statusKeys= ["Deferred", "Scuttled"];
  var statusPreambleFiller= " due to: ";
  var statusPreamble= statusAction + statusPreambleFiller;
  
  if (!statusPrior.includes(statusPreamble))
  {
    // Add preamble since it is not included yet 
    statusPrior= statusPreamble + statusPrior;
  }
  else
  {
    // Update an existing preamble
    for (var statusKey in statusKeys)
    {
      // Check if another action preserved status prior to this attempt
      if (statusPrior.includes(statusKey + statusPreambleFiller))
      {
        if (statusKey != statusAction)
        {
          // Replace the previous, non-matching  preserving action with the current one
          statusPrior.replace(statusKey + statusPreambleFiller, statusPreamble);
        }
        
        break;
      }
    }
  }
  
  UpdateStatus(parameters, statusPrior);
};


/**
 * UpdateAlertStatus()
 *
 * Update status and time of alert
 */
function UpdateAlertStatus(parameters, price, alert, time)
{
  SetValueByName(parameters["id"], "statusAlert", alert, parameters["verbose"]);
  SetValueByName(parameters["id"], "statusAlertTime", DateToLocaleString(time), parameters["verbose"]);
  SetValueByName(parameters["id"], "priceAlertLast", price, parameters["verbose"]);
};


/**
 * UpdateRunStamps()
 *
 * Update status and time of status
 */
function UpdateRunStamps(parameters)
{
  var success= false;
  
   if (success= SetValueByName(parameters["id"], "statusRunCurrent", parameters["scriptTime"].toFixed(0), parameters["verbose"]))
   {
     SetValueByName(parameters["id"], "statusRunCurrentTime", DateToLocaleString(parameters["scriptTime"]), parameters["verbose"]);
     SetValueByName(parameters["id"], "statusRunPrevious", parameters["statusRunCurrent"], parameters["verbose"]);
     SetValueByName(parameters["id"], "statusRunPreviousTime", DateToLocaleString(parameters["statusRunCurrent"]), parameters["verbose"]);
   }
  
  return success;
};


/**
 * UpdatePreviousRegressionSlope()
 *
 * Preserve previous regression slope
 */
function UpdatePreviousRegressionSlope(parameters)
{
  var success= false;
  
  UpdateStatus(parameters, "Saving previous regression slope...");
  
  if (success= SetValueByName(parameters["id"], "priceRegressionSlopePrevious", parameters["priceRegressionSlope"], parameters["verbose"]))
  {
    SetValueByName(parameters["id"], "priceRegressionSlopePreviousTime", DateToLocaleString(), parameters["verbose"]);
  }
  
  return success;
};


/**
 * UpdateLastPrice()
 *
 * Preserve latest price
 */
function UpdateLastPrice(parameters, price)
{
  var success= false;
  
  if (success= SetValueByName(parameters["id"], "priceLast", price, parameters["verbose"]))
  {
    parameters["priceLast"]= price;
    SetValueByName(parameters["id"], "priceLastTime", DateToLocaleString(), parameters["verbose"]);
  }
  
  return success;
};


/**
 * UpdateURL()
 *
 * Preserve latest query URL
 */
function UpdateURL(parameters, url)
{
  SetValueByName(parameters["id"], "comedURL", url, parameters["verbose"]);
  SetValueByName(parameters["id"], "comedURLTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * UpdateMissingPricesAlertStamp()
 *
 * Preserce the step of the last time we alerted about missing (delayed) prices
 */
function UpdateMissingPricesAlertStamp(parameters)
{
  SetValueByName(parameters["id"], "noNewPricesAlert", parameters["scriptTime"], parameters["verbose"]);
  SetValueByName(parameters["id"], "noNewPricesAlertTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * ClearMissingPricesAlertStamp()
 *
 * Preserce the step of the last time we alerted about missing (delayed) prices
 */
function ClearMissingPricesAlertStamp(parameters)
{
  SetValueByName(parameters["id"], "noNewPricesAlert", "", parameters["verbose"]);
  SetValueByName(parameters["id"], "noNewPricesAlertTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * PrepareToCommit()
 *
 * Prepare for writing and alerting
 */
function PrepareToCommit(parameters)
{
  var success= false;
  var activity= "Preparing to commit";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"]= activity.toLowerCase();
  
  if (success= IsSupreme(parameters))
  {
    // This is still the latest run: grab and preserve latest regression slope and moving average values 
    if (success= SetSemaphore(parameters))
    {
      // No conflicting runs -- proceed
      if (success= UpdateRunStamps(parameters))
      {
        success= UpdatePreviousRegressionSlope(parameters);
      }
    }
  }
  else
  {
    Superseded(parameters, "PrepareToCommit", activity); 
  }
  
  return success;
};


/**
 * SavePrices()
 *
 * Save obtained prices
 */
function SavePrices(parameters, prices)
{
  var success= false;
  var activity= "Saving prices";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"]= activity.toLowerCase();
  
  if (success= IsSupreme(parameters))
  {
    // This is still the latest run: write obtained prices to history table 
    prices[prices.length-1][parameters["indexStamp"]]= parameters["scriptTime"].toFixed(0);
    prices[prices.length-1][parameters["indexStampTime"]]= DateToLocaleString();
    success= SaveSnapshot(parameters["id"], parameters["comedSheetPrices"], prices, parameters["verbose"]);
  }
  else
  {
    Superseded(parameters, "SavePrices", activity); 
  }
  
  return success;
};


/**
 * UpdateComputedValues()
 *
 * Update trend and moving average
 */
function UpdateComputedValues(parameters)
{
  var success= false;
  var onlyIfBlank= true;
  var priceMovingAverage= null;
  var priceTrend= null;
  var activity= "Updating computed values";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"]= activity.toLowerCase();
  
  if (success= IsSupreme(parameters))
  {
    // This is still the latest run: update freshly recomputed values
    parameters["activity"]= "updating moving average";
    
    priceMovingAverage= GetValueByName(parameters["id"], "priceMovingAverage", parameters["verbose"], parameters["confirmNumbers"], parameters["confirmNumbersLimit"]);
    if (priceMovingAverage != null)
    {
      parameters["priceMovingAverage"]= priceMovingAverage;
      success= UpdateSnapshotCell(parameters["id"], parameters["comedSheetPrices"], parameters["indexMovingAverage"] + 1, priceMovingAverage,
                                  onlyIfBlank, parameters["verbose"]);
    }
    else
    {
      success= false;
      Logger.log("[UpdateComputedValues] Could not obtain updated Moving Average!");
    }
    
    if (success)
    {
      // Success so far: proceed... 
      parameters["activity"]= "updating trend";
      
      priceTrend= GetValueByName(parameters["id"], "priceRegressionSlope", parameters["verbose"], parameters["confirmNumbers"], parameters["confirmNumbersLimit"]);
      if (priceTrend != null)
      {
        parameters["priceRegressionSlope"]= priceTrend;
        success= UpdateSnapshotCell(parameters["id"], parameters["comedSheetPrices"], parameters["indexTrend"] + 1, priceTrend, onlyIfBlank, parameters["verbose"]);
      }
      else
      {
        success= false;
        Logger.log("[UpdateComputedValues] Could not obtain updated Regression Coefficient.");
      }
    }
  }
  else
  {
    Superseded(parameters, "SavePrices", activity); 
  }
  
  return success;
};


/**
 * Notify()
 *
 * Notify via alerts, if triggered by parameters
 */
function Notify(parameters)
{
  var success= false;
  var onlyIfBlank= true;
  var statusMessage= null;
  var activity= "Preparing to notify";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"]= activity.toLowerCase();
  
  if (success= IsSupreme(parameters))
  {
    // This is still the latest run: notify, if necessary
    statusMessage= SendPriceAlert(parameters);
    if (statusMessage.length > 0)
    {
      UpdateStatus(parameters, statusMessage);
      success= UpdateSnapshotCell(parameters["id"], parameters["comedSheetPrices"], parameters["indexAlert"] + 1, statusMessage, onlyIfBlank, parameters["verbose"]);
    }
    else
    {
      Logger.log("[Notify] Received empty status report <%s> after trying to send price alert.", statusMessage);
    }
  }
  else
  {
    Superseded(parameters, "Notify", activity); 
  }
  
  return success;
};


/**
 * ConvertMillisecondsToMinutes()
 *
 * Return elapsed time in minutes, converted from milliseconds
 */
function ConvertMillisecondsToMinutes(milliseconds)
{
  return (milliseconds / 60 / 1000).toFixed(2);
};