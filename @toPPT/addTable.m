function addTable(ppt, title, text, varargin)
    %% addTable(title, text, varargin)
    % titles is the titles of the columns of the table
    % text is the text to be generated, using css tags such as
    % <s font-family:Times New Roman> text </s>
    % <s font-size:22> text </s>
    % <s font-family:Times New Roman;font-size:22> text </s>
    % all allowable tags:
    %   bg: red, blue, green, orange, purple...
    %       <a href="matlab:help rgb">additional colors</a>
    %   font-family: Times New Roman, Ariel, Comis Sans MS
    %       any allowable font names
    %   font-size: 1 to 100
    %   color: same allowable values as bg
    %   href: link value (255,@,slide#)
    %
    % additioanl tags that can be used:
    %   <b> </b> to bold text
    %   <u> </u> to underline text
    %   <i> </i> to italicize text
    %
    % Allowable additional variables to pass in
    % Location - [x%,y%] where this is from the NW corner of the
    %   slide, defaults to [5, 25]
    % Size - [x%,y%] gives the size of the text box to be created
    %   defaults to [90, 70]
    % Column - [x% array] sets the width of each column based on the array
    %   passed, will normalize if the sum is not 100
    % Merge - {cell array} meges cells in the table. The cells gives an 2x2
    %   matrix that gives the column and row to be merged formatted 
    %   [x1,y1;x2,y2] 
    %
    % Notes on the text:
    %   Needs to be a cell array each new line is a new cell
    %   Each line without tags will default to Times New Roman,
    %       20 Pt font
    %
    
    [rows,cols] = size(text);
    
    location = [5,25];
    %even spacing columns
    columnarray= ones(1,cols)*100/cols;
    doMerge = 0;
    MergeArray = {};
    boxSize = [90, 70];
    for a = 1:2:nargin-3
        %case statement to parse out values
        switch upper(varargin{a})
            case 'COLUMN'
                columnarray = varargin{a+1};
                if(~isnumeric(columnarray))
                    error('Column needs to be a numeric array passed in')
                elseif(length(columnarray) > cols)
                    warning('More values in the column than there are columns in the text')
                    columnarray = columnarray(1:cols);
                end
                if(sum(columnarray) ~= 100)
                    %rescale the column array
                    columnarray = columnarray*100/sum(columnarray);
                end
            case 'LOCATION'
                location = varargin{a+1};
                if(numel(location) ~= 2)
                    error('Location must be a 2 value array')
                elseif(~isnumeric(location))
                    error('Location must be passed as a numeric value')
                elseif(sum(location > 100 | location < 0))
                    error('Location must be values between 0 and 100')
                end
            case 'MERGE'
                doMerge = 1;
                MergeArray = varargin{a+1};
                if(~iscell(MergeArray))
                    error('Merge Array needs to be passed in a cell array of 2x2 matricies')
                end
            case 'SIZE'
                boxSize = varargin{a+1};
                if(numel(boxSize) ~= 2)
                    error('Size must be a 2 value array')
                elseif(~isnumeric(boxSize))
                    error('Size must be passed as a numeric value')
                elseif(sum(boxSize > 100 | boxSize < 0))
                    error('Size must be values between 0 and 100')
                end
            otherwise
                warning('Invalid input detected')
        end
    end
    
    %get the current slide shape to calculate the location and
    %width in points of the text box
    slideHeight = ppt.presentation.PageSetup.SlideHeight;
    slideWidth = ppt.presentation.PageSetup.SlideWidth;

    %Create the Textbox
    box = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count).Shapes.AddTable(...
        rows+1,cols,slideWidth*location(1)/100,slideHeight*location(2)/100,...
        slideWidth*boxSize(1)/100,slideHeight*boxSize(2)/100);
    
    %Set tiles in first row and size column widths
    for a = 1:cols
        ppt.addFormattedText(box.Table.Cell(1,a).Shape.TextFrame.TextRange,title{a},1);
        %must be set in pixles because... reasons
        box.Table.Columns.Item(a).Width = slideWidth*boxSize(1)/100*columnarray(a)/100;
    end
    
    %go through each row and populate the cells
    for a = 1:rows
        for b = 1:cols
            ppt.addFormattedText(box.Table.Cell(a+1,b).Shape.TextFrame.TextRange,text{a,b},1);
        end
    end
    
    try
        if(doMerge)
            for a = 1:length(MergeArray)
                box.Table.Cell(MergeArray{a}(1,2),MergeArray{a}(1,1)).Merge(...
                    box.Table.Cell(MergeArray{a}(2,2),MergeArray{a}(2,1)));
            end
        end
    catch
        error('Merging formatting failed')
    end
    
end