classdef toPPT < handle
    %toWord Summary of this class goes here
    %   Detailed explanation goes here
    
    properties (Hidden)
        activeXCom = [];
        presentation = [];
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
        
        function addText(ppt, text, varargin)
            %% addText(text, varargin)
            % text is the text to be generated, using css tags such as
            % <s font-family:Times New Roman> text </s>
            % <s font-size:22> text </s>
            % <s font-family:Times New Roman;font-size:22> text </s>
            % all allowable tags:
            % bg: red, blue, green, orange, purple... 
            %       <a href="matlab:help rgb">additional colors</a>
            % font-family: Times New Roman, Ariel, Comis Sans MS
            %       any allowable font names
            % font-size: 1 to 100
            % href: link value (255,@,slide#)
            %
            % additioanl tags that can be used:
            % <b> </b> to bold text
            % <u> </u> to underline text
            % <i> </i> to italicize text
            %
            % Allowable additional variables to pass in
            % Location - [x%,y%] where this is from the NW corner of the
            %   slide, defaults to [5, 20]
            % Size - [x%,y%] gives the size of the text box to be created
            %   defaults to [90, 50]
            % Bullets - can be used to give each line of text a bullet,
            %   needs to be a character. Defaults to no bullets.
            % Indent - can be used to indent the line of text to be paired
            %   with Bullets
            
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
    end
    
end