classdef toPPT < handle
    %toWord Summary of this class goes here
    %   Detailed explanation goes here
    
    properties (Hidden)
        activeXCom = [];
        presentation = [];
        currentSlide = [];
    end
    
    methods
        
        function ppt = toPPT(varargin)
            %constructs and opens a port to Word
            ppt.activeXCom = actxserver('PowerPoint.Application');
        end
        
        function Open(ppt, file)
            %% Open(file)
            % Opens an existing power point presentation file.
            try
                ppt.activeXCom.Presentations.Open(file);
                ppt.presentation = wordClass.activeXCom.ActivePresentation;
            catch
                error('Seems the file does not exist')
            end
        end
        
        function NewTemplate(ppt)
            %% NewTemplate
            % Creates a new power point presentation from the Template.pptx
            % file in the toPPT class folder
            try
                out = which('toPPT');
                out = out(1:find(out == '\',1,'last'));
                ppt.activeXCom.Presentations.Open([out 'Template.pptx']);
                ppt.presentation = ppt.activeXCom.ActivePresentation;
            catch
                error('Seems the file does not exist')
            end
        end
        
        function New(ppt)
            %% New
            % Creates a new blank power point presentation
            try
                ppt.presentation = ppt.activeXCom.Presentation.Add;
            catch
                error('unable to open new document')
            end
        end
        
        function addSection(ppt, SectionTitle)
            %% addSection(SectionTitle)
            % adds a new section title to the presentation before the
            % current slide
            
            try
                %assume last slide is current slide for now
                ppt.presentation.SectionProperties.AddSection(...
                    ppt.presentation.Slides.Count,SectionTitle)
            catch
                error('Slide title was not valid')
            end
        end
        
        function newSlide(ppt,slideIndex)
            %% newSlide(slideIndex)
            % adds a new slide to the power point presentation
            %if slideIndex is empty, will add to the end of the
            %presentation, if slideIndex is not will add it there
            
            try
                if(nargin < 2)
                    %note:CustomLayouts.Item(11) is the "first" slide in
                    %the slide master layout, 1 is a title slide.  Should
                    %enable selection of the slide at some point
                    ppt.currentSlide = ppt.presentation.Slides.AddSlide(ppt.presentation.Slides.Count+1,...
                        ppt.presentation.SlideMaster.CustomLayouts.Item(11));
                else
                    ppt.currentSlide = ppt.presentation.Slides.AddSlide(slideIndex,...
                        ppt.presentation.SlideMaster.CustomLayouts.Item(11));
                end
            catch
                error('adding slide failed')
            end
        end
        
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
            box = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count).Shapes.AddTextbox('msoTextOrientationHorizontal',...
                slideWidth*location(1)/100,slideHeight*location(2)/100,slideWidth*boxSize(1)/100,slideHeight*boxSize(2)/100);
            
            previousLineCount = 0;
            for a = 1:textLines
                %add the text to the box and format it correctly
                addFormattedText(box.TextFrame.TextRange,text{a});
                %set the bullet of the line
                if(setBullets)
                    if(Bullets(a) == 0)
                        box.TextFrame.TextRange.Lines(previousLineCount+1).ParagraphFormat.Bullet.Type = 'ppBulletNone';
                    else
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
        
    end
    
    methods(Hidden)
        
        %% hide all the base class function from tab complete list
        function addlistener(fte, property, eventname, callback)
            addlistener@addlistener(fte, property, eventname, callback)
        end

        function eq(fte, A,B)
            eq@eq(fte, A,B)
        end

        function findobj(fte, varargin)
            findobj@handle(fte, varargin);
        end

        function findprop(fte, name)
            findprop@handle(fte,name);
        end

        function ge(fte, A,B)
            ge@ge(fte, A,B);
        end

        function gt(fte, A,B)
            gt@gt(fte,A,B);
        end

        function le(fte, A,B)
            le@le(fte,A,B);
        end

        function lt(fte, A,B)
            lt@lt(fte,A,B);
        end

        function ne(fte, A,B)
            ne@ne(fte,A,B);
        end

        function notify(fte, varargin)
            notify@handle(fte, varargin);
        end

        function delete( ppt )
            %% force cleanup
            release(ppt.activeXCom);
            delete(ppt.activeXCom);
            delete@handle(ppt);
        end
        
        function addFormattedText(ppt, textRange, text)
            %% interprets the text and the new text to the slide
            
            
        end
    end
    
end