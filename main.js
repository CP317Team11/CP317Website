function loadExcelFile(filePath, sheetName) {
    console.log(`Attempting to load Excel file from: ${filePath}`);
    fetch(filePath)
        .then(response => {
            if (!response.ok) {
                throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });

           
            if (!workbook.SheetNames.includes(sheetName)) {
                throw new Error(`Sheet "${sheetName}" not found in the Excel file.`);
            }

           
            const sheet = workbook.Sheets[sheetName];
            console.log(`Using sheet: ${sheetName}`);

            const excelData = XLSX.utils.sheet_to_json(sheet);

            
            excelData.forEach(event => {
                if (typeof event.Date === 'number') {
                    event.Date = convertExcelDate(event.Date); 
                }
            });
            console.log('Parsed events with corrected dates:', excelData);

            const calendarEvents = transformExcelDataToCalendarEvents(excelData);
            console.log('Transformed Calendar Events:', calendarEvents);

   
            initializeCalendar(calendarEvents);
        })
        .catch(error => console.error('Error loading Excel file:', error));
}


function convertExcelDate(serial) {
    const excelEpoch = new Date(1900, 0, 1);
    return new Date(excelEpoch.getTime() + (serial - 1) * 86400000).toISOString().split('T')[0];
}


function displayExcelContents(events) {
    const displayDiv = document.createElement("div");
    displayDiv.id = "excel-contents";
    const title = document.createElement("h2");
    title.innerText = "Contents of the Excel File:";
    displayDiv.appendChild(title);

    events.forEach((event, index) => {
        const eventDetails = document.createElement("p");
        eventDetails.innerText = `Row ${index + 1}: Organization - ${event["Organization"]}, Event Name - ${event["Event Name"]}, Location - ${event["Location"]}, Date - ${event["Date"]}`;
        displayDiv.appendChild(eventDetails);
    });

    document.body.appendChild(displayDiv);
}


function initializeCalendar(events) {
    $('#calendar').fullCalendar({
        defaultView: 'month',
        editable: false,
        events: events, 
        eventClick: function(event) {
            alert(`Event: ${event.title}\nDate: ${event.start.format('YYYY-MM-DD')}`);
        }
    });
}

function transformExcelDataToCalendarEvents(events) {
    return events.map(event => ({
        title: event["Event Name"],
        start: event["Date"], 
        description: `${event["Organization"]}, ${event["Location"]}`
    }));
}


$(document).ready(function () {
    console.log("Document is ready.");


    if (window.location.pathname.includes("wilfrid-laurier.html")) {
        console.log("Loading WLU Events sheet...");
        loadExcelFile('eventsdatabase.xlsx', 'WLU Events');
    } else if (window.location.pathname.includes("waterloo.html")) {
        console.log("Loading UW Events sheet...");
        loadExcelFile('eventsdatabase.xlsx', 'UW Events');
    } else if (window.location.pathname.includes("mcmaster.html")) {
        console.log("Loading McMaster Events sheet...");
        loadExcelFile('eventsdatabase.xlsx', 'MCMASTER Events');
    } else if (window.location.pathname.includes("uoft.html")) {
        console.log("Loading UofT Events sheet...");
        loadExcelFile('eventsdatabase.xlsx', 'UOFT Events');
    } else if (window.location.pathname.includes("yorkuni.html")) {
        console.log("Loading York Events sheet...");
        loadExcelFile('eventsdatabase.xlsx', 'YORK Events');
    } else {
        console.log("No matching page for Excel sheet loading.");
    }
});




function initializeCalendar(events) {
    $('#calendar').fullCalendar({
        defaultView: 'month',
        editable: false,
        events: events, 
        eventClick: function(event) {
           
            $('#eventTitle').text(event.title);
            $('#eventDate').text(moment(event.start).format('MMMM Do YYYY'));
            $('#eventDescription').text(event.description || 'No additional details available.');
            

            $('#eventModal').fadeIn();
        }
    });
}


$('.close, .close-btn').click(function() {
    $('#eventModal').fadeOut();
});


$(window).click(function(event) {
    if ($(event.target).is('#eventModal')) {
        $('#eventModal').fadeOut();
    }
});

const scriptURL = 'https://script.google.com/macros/s/AKfycbwuKAhq2tBPPQwCCbHZfDR2KcyPPy_0vL7XCe9I_cHeALe9hcPJWf_QUCzzlh63E5vtHA/exec'

    const form = document.forms['signup-form']
    
    form.addEventListener('submit', e => {
      
      e.preventDefault()
      
      fetch(scriptURL, { method: 'POST', body: new FormData(form)})
      .then(response => alert("Thank you! Form is submitted" ))
      .then(() => { window.location.reload(); })
      .catch(error => console.error('Error!', error.message))
    })


