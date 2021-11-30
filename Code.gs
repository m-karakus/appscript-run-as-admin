function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GetValues')
        .addItem('get current', 'run')
        .addToUi();
}

function run(){
  var url = "https://script.google.com/macros/s/......../exec"
  var result = UrlFetchApp.fetch(url);
}

function doGet() {
    var ss = SpreadsheetApp.openById('143GQYzA4Y9Gg2Lmk-J9nuMNI6FH3H-sUbi4v463saa8');
    var sh0 = ss.getSheetByName("Deploy");
    var row = 2;
    var col = 1;

    var value = (sh0.getRange(row, col).getValue());
    var arr = [];
    for (i = -6; i < +8; i++) {
        var result = new Date(value);
        result.setDate(value.getDate() + i);
        result = result.toISOString().slice(0, 10).replace(/-/g, "-");
        arr.push(result);
    };


    var result = new Date(value);
    result.setDate(value.getDate() - 14);
    result = result.toISOString().slice(0, 10).replace(/-/g, "-");
    arr.push(result);

    var result = new Date(value);
    result.setDate(value.getDate() + 16);
    result = result.toISOString().slice(0, 10).replace(/-/g, "-");
    arr.push(result);

    var result = new Date(value);
    result.setDate(value.getDate() - 29);
    result = result.toISOString().slice(0, 10).replace(/-/g, "-");
    arr.push(result);

    var result = new Date(value);
    result.setDate(value.getDate() + 31);
    result = result.toISOString().slice(0, 10).replace(/-/g, "-");
    arr.push(result);


    var quotedAndCommaSeparated = "'" + arr.join("','") + "'";
    //console.log(quotedAndCommaSeparated);

    query = "with t as (select * from `ksbigquery.KS.FollowingMetricsReport` as f  UNION ALL select 'Google Entrance' as Type_,date(e.date) as Date_, sum(e.entrances) as Count_	  from `ksbigquery.KS.vw_KS_GA_Entrances` as e  where date(e.date) > '2021-07-01'group by date(e.date) UNION ALL select  'Superb - MHO by Auto' as Type_,date(a.PostedOn) as Date_, count(*) as Count_	 from `ksbigquery.KS.Answer` as a where a.AutoSelectedBestAnswer !='0' and date(a.PostedOn) > '2021-07-01' group by date(a.PostedOn)) select * from t as f PIVOT ( sum(Count_) as n FOR Date_ in (" + quotedAndCommaSeparated + "))"
    console.log(query)

    var projectId = 'ksbigquery';
    var request = {
        query: query,
        useLegacySql: false
    };
    var queryResults = BigQuery.Jobs.query(request, projectId);
    var jobId = queryResults.jobReference.jobId;

    // Check on status of the Query Job.
    var sleepTimeMs = 500;
    while (!queryResults.jobComplete) {
        Utilities.sleep(sleepTimeMs);
        sleepTimeMs *= 2;
        queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
    }

    // Get all the rows of results.
    var rows = queryResults.rows;
    while (queryResults.pageToken) {
        queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
            pageToken: queryResults.pageToken
        });
        rows = rows.concat(queryResults.rows);
    }

    if (rows) {
        // Append the headers.
        var headers = queryResults.schema.fields.map(function (field) {
            return field.name;
        });
        //sheet.appendRow(headers);
        // Append the results.
        var data = new Array(rows.length);
        for (var i = 0; i < rows.length; i++) {
            var cols = rows[i].f;
            data[i] = new Array(cols.length);
            for (var j = 0; j < cols.length; j++) {
                data[i][j] = cols[j].v;
            }
        }
        sh0.getRange(5, 1, rows.length, headers.length).setValues(data);

        Logger.log('Results spreadsheet created: %s',
            ss.getUrl());
    } else {
        Logger.log('No rows returned.');
    }
}