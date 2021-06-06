% Close any open figures
close all;
% Messages for the user 
disp('Click and drag the mouse whereever in the figure.');
disp('--On release it will open up a spreadsheet containing plotted data');
disp('--WARNING: Kills any excel processes');
% Create the figure background and limit the x,y coords to prevent the
% graph from auto adjusting
figure();
xlim([0 1]);
ylim([0 1]);
% Set hold on to preserve all plotted points
hold on;

% Initalizes global record variable. 
% Record is used to track whether or not the app should record user's mouse
setGlobalRecord(false);

% Window functions using the active window and using callbacks on events
set(gcf, 'WindowButtonDownFcn', @mouseDown);
set(gcf, 'WindowButtonMotionFcn', @mouseMove);
set(gcf, 'WindowButtonUpFcn', @mouseUp);

% On mouse down. Record = true, Time = epoch time.
% Time is used to track start time of mouse down and sequential mouse move times. 
function mouseDown (objectHandle, ~)
    % Start recording and the timer
    setGlobalRecord(true);
    setGlobalStartTime(posixtime(datetime('now')));
end

% Sets record = false to stop recording mouse. 
% Finds our figure and extracts the data from it.
% Passes said data to writeExcel function.
% Closes all figures on mouse up (mouse release)
function mouseUp (objectHandle, ~)
    % Stop recording data
    setGlobalRecord(false);
    
    fig = findobj(gcf, 'Type', 'line');
    x = get(fig, 'Xdata');
    y = get(fig, 'Ydata'); 
    z = get(fig, 'Zdata'); 
    filename = 'data.xlsx';
    writeExcel(filename, x, y, z);
    % Closes all figures
    close all;
end

% Plots mouse movement if record is not false(0)
function mouseMove (objectHandle, ~)
    r = getGlobalRecord;
    if r == 0 % Record is false;
        return
    end
    
    plotPoint(objectHandle, 'b.', 8);
end

function plotPoint(objectHandle, markerStyle, markerSize)
    axesHandle  = get(objectHandle, 'CurrentAxes');
    point = get(axesHandle, 'CurrentPoint'); 
    point = point(1,1:2);
    x = point(1);
    y = point(2);
    t = getGlobalStartTime;
    % Substracts curtime from start time to get a sequential time on mouse
    % movements
    z = posixtime(datetime('now')) - t;
    plot3(x, y, z, markerStyle, 'MarkerSize', markerSize);
end

function points = calculateExcelData(x ,y, z)
    x = flip(x);
    y = flip(y);
    z = flip(z);
    
    X = x;
    Y = y;
    time = z;

    points = [X,Y, time];
    
end

% Closes any existing excel processes. 
% Deletes filename to wipe the data.
% Rewrites filename with x,y,z as individual columns.
% Flips the x,y,z data to be sequential in time.
% Inserts a scatterplot with x,y which displays what was drawn in this
% application.
% Then opens the spreadsheet containing x,y,z and the scatterplot.
function writeExcel(filename, x, y, z)

    % Kill any excels to delete our target data file
    % Soft errors if not found.
    system('taskkill /F /IM EXCEL.EXE');
    % Delete data file
    if isfile(filename)
      delete(filename);
    end
    points = calculateExcelData(x, y, z);

    % Write data to data file
    xlswrite(filename, points);

    % Open excel app
    excel = actxserver('Excel.Application');
    % Get project working dir and get full path to data file
    dir = pwd;
    fileFullPath = strcat(dir,'\',filename);
    % Open workbook using data full file path
    wb = excel.Workbooks.Open(fileFullPath);
    % Find sheet1 
    workSheets = wb.Sheets;  
    myWorkSheet = workSheets.Item('Sheet1');  

    % create an object of ChartObjects
    myChartObject = myWorkSheet.ChartObjects.Add(100, 30, 400, 250);  
    % create an object of Chart. 
    myPlots = myChartObject.Chart; 
    myPlots.HasTitle = true;
    myPlots.ChartTitle.Text = 'Raw Data';
    % create an object of SeriesCollection (XY plot for the raw data)
    line = myPlots.SeriesCollection.NewSeries;  
    % Set X,Y Data for line graph
    myPlots.SeriesCollection(1).XValue = myWorkSheet.Range('A1:A10000'); % Range A:A is really slow for some reason
    myPlots.SeriesCollection(1).Values = myWorkSheet.Range('B1:B10000');
    % Set chart type and name
    line.ChartType = 'xlXYScatterLines';
    line.Name = 'Points';
    % Save the workbook and quit.
    excel.ActiveWorkbook.Save;
    excel.Quit;
    % Reopen sheet
    % excel.Visible = true wouldnt prioritize window so this is the
    % next best thing.
    if isfile(filename)
       winopen(fileFullPath);
    end
end

function setGlobalRecord(val)
    global Record
    Record = val;
end

function r = getGlobalRecord
    global Record
    r = Record;
end

function setGlobalStartTime(val)
    global StartTime
    StartTime = val;
end

function r = getGlobalStartTime
    global StartTime
    r = StartTime;
end
