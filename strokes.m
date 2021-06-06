disp('Outputs critical data into an excel sheet. Requires data.xlsx to be in the same directory.');
% Data file directory
dataFileName = 'data.xlsx';
dir = pwd;
dataFileFullPath = strcat(dir,'\',dataFileName);
% Import data from excel and get the important stuff
data = xlsread(dataFileFullPath);
x = data(:,1);
y = data(:,2);
z = data(:,3);

filename = 'criticals.xlsx';
criticalData = calculateCriticals(x,y,z);

writeExcel(filename, criticalData);

function criticalData = calculateCriticals(x,y,z)

    % Format our data as cells 
    x = num2cell(x);
    y = num2cell(y);
    time = num2cell(z);

    % Allocate space for our upcoming cells
    xDiff{length(x)-1,1} = [];
    yDiff{length(y)-1,1} = [];
    norm{length(y)-1,1} = [];
    smooth{length(y)-1,1} = [];

    % Calculate difference x
    for i = length(x):-1:2
       xDiff{i-1} = x{i} - x{i-1};
    end
    xDiff{end+1} = xDiff{end};

    % Calculate difference y
    for i = length(y):-1:2
       yDiff{i-1} = y{i} - y{i-1};
    end
    yDiff{end+1} = yDiff{end};

    % Calculate norm
    for i = length(x):-1:1
       norm{i} = sqrt(xDiff{i}^2  + yDiff{i}^2);
    end

    % Get avg value of the norm
    avg = mean(cell2mat(norm));
    % Skips prevents similar critical points from being plotted
    skips = 0;
    % Allocate space for critical cells
    criticalX = cell(2, 1);
    criticalY = cell(2, 1);
    criticalZ = cell(2, 1);
  
    % Ensure the final data points are in the criticals
    criticalX{end} = x{end};
    criticalY{end} = y{end};
    criticalZ{end} = time{end};  
    
    % Calculate smooth f
    for i = length(x):-1:5
        smooth{i} = norm{i} + norm{i-1} + norm{i-2} + norm{i-3} + norm{i-4};
        smooth{i} = smooth{i}/5;

        % if the smoothing value is lower than the avg smooth value then
        % its probably important
        if smooth{i} < avg
            % if a point has been plotted recently (skips > 0) then we dont
            % bother plotting it
            if (skips > 0)
                skips = skips - 1;
            else 
                skips = 3;
                criticalX{end + 1} = x{i};
                criticalY{end + 1} = y{i};
                criticalZ{end + 1} = time{i};                   
            end 
        end
    end
    % Ensure the first data points are in the criticals
    criticalX{end} = x{1};
    criticalY{end} = y{1};
    criticalZ{end} = time{1};     
    % Assign the critical data into an array and flip it so that its sorted
    % properly in the excel sheet later
    criticalData = [flip(criticalX), flip(criticalY), flip(criticalZ)];

end

function writeExcel(filename, data)

    % Kill any excels to delete our target data file
    % Soft errors if not found.
    system('taskkill /F /IM EXCEL.EXE');
    % Delete data file
    if isfile(filename)
      delete(filename);
    end

    % Write data to data file
    xlswrite(filename, data);

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