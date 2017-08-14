function addText(ppt, text, varargin)
    %% addText(text, varargin)
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
    %   shref: link value (255,@,slide#)
    %   href: link value (http://www.google.com)
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
    % Bullets - can be used to give each line of text a bullet,
    %   needs to be a character. Defaults to no bullets.
    % Indent - can be used to indent the line of text to be paired
    %   with Bullets
    %
    % Notes on the text:
    %   Needs to be a cell array each new line is a new cell
    %   Each line without tags will default to Times New Roman,
    %       20 Pt font
    %

    setBullets = 0;
    location = [5,25];
    setIndents = 0;
    textLines = length(text);
    boxSize = [90, 70];
    for a = 1:2:nargin-2
        %case statement to parse out values
        switch upper(varargin{a})
            case 'BULLETS'
                setBullets = 1;
                Bullets = varargin{a+1};
                if(length(Bullets) < textLines)
                    error('Bullets passed in does not cover all the lines, try again with at least the right number of lines')
                elseif(length(Bullets) > textLines)
                    warning('Number of bullets passed in is greater than the number of text lines, you may want to check')
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
            case 'INDENT'
                setIndents = 1;
                Indent = varargin{a+1};
                if(length(Indent) < textLines)
                    error('Indent passed in does not cover all the lines, try again with at least the right number of lines')
                elseif(length(Indent) > textLines)
                    warning('Number of Indent passed in is greater than the number of text lines, you may want to check')
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
    box = ppt.presentation.Slides.Item(ppt.currentSlide.SlideNumber).Shapes.AddTextbox('msoTextOrientationHorizontal',...
        slideWidth*location(1)/100,slideHeight*location(2)/100,slideWidth*boxSize(1)/100,slideHeight*boxSize(2)/100);

    previousLineCount = 0;
    for a = 1:textLines
        %add the text to the box and format it correctly
        ppt.addFormattedText(box.TextFrame.TextRange,text{a},a==textLines);
        %set the bullet of the line
        if(setBullets)
            if(Bullets(a) == 0)
                box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Type = 'ppBulletNone';
            elseif(Bullets(a) == 1)
                %1 is a -
                box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Character = 45;
            elseif(Bullets(a) == 2)
                %2 is a dot
                box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Character = 8226;
            elseif(Bullets(a) == 3)
                %3 is a box
                box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Character = 9632;
            else
                %everything else is just the character
                box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Character = Bullets(a);
            end
        end
        if(setIndents)
            %indent all the lines just added
            for b = previousLineCount+1:box.TextFrame.TextRange.Lines.Count
                box.TextFrame.TextRange.Lines(b).IndentLevel = Indent(a);
            end
        end
        previousLineCount = box.TextFrame.TextRange.Lines.Count;
    end

end