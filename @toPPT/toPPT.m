classdef toPPT < handle
    %toPPT Summary of this class goes here
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
        
        function open(ppt, file)
            %% Open(file)
            % Opens an existing power point presentation file.
            try
                ppt.activeXCom.Presentations.Open(file);
                ppt.presentation = ppt.activeXCom.ActivePresentation;
                ppt.currentSlide = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count);
                %select the first slide to allow setting of the slide
                ppt.currentSlide.Select;
            catch
                error('Seems the file does not exist')
            end
        end
        
        function newTemplate(ppt)
            %% newTemplate
            % Creates a new power point presentation from the Template.pptx
            % file in the toPPT class folder
            try
                out = which('toPPT');
                out = out(1:find(out == '\',1,'last'));
                ppt.activeXCom.Presentations.Open([out 'Template.pptx']);
                ppt.presentation = ppt.activeXCom.ActivePresentation;
                ppt.currentSlide = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count);
                %select the first slide to allow setting of the slide
                ppt.currentSlide.Select;
            catch
                error('Seems the file does not exist')
            end
        end
        
        function new(ppt)
            %% new
            % Creates a new blank power point presentation
            try
                ppt.presentation = ppt.activeXCom.Presentation.Add;
                ppt.currentSlide = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count);
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
                ppt.presentation.SectionProperties.AddBeforeSlide(...
                    ppt.presentation.Slides.Count,SectionTitle);
            catch
                error('Slide title was not valid')
            end
        end
        
        function setTitle(ppt, Title)
            %% setTitle(Title)
            % sets the title of the current slide
            if(ischar(Title))
                try
                    if(strcmp(ppt.currentSlide.Shapes.HasTitle,'msoFalse'))
                        ppt.currentSlide.Shapes.AddTitle.TextFrame.TextRange.Text = Title;
                    else
                        ppt.currentSlide.Shapes.Title.TextFrame.TextRange.Text = Title;
                    end
                catch
                    error('Adding slide title failed')
                end
            end
        end
        
        function newSlide(ppt,slideIndex)
            %% newSlide(slideIndex)
            % adds a new slide to the power point presentation
            % if slideIndex is empty, will add to the end of the
            % presentation, if slideIndex is not will add it there
            
            try
                if(nargin < 2)
                    %note:CustomLayouts.Item(6) is the "blank" slide in
                    %the slide master layout, 1 is a title slide.  Should
                    %enable selection of the slide at some point
                    %ppt.currentSlide = ppt.presentation.Slides.AddSlide(ppt.presentation.Slides.Count+1,...
                    %    ppt.presentation.SlideMaster.CustomLayouts.Item(6));
                    ppt.currentSlide = invoke(ppt.presentation.Slides,'Add',...
                        ppt.presentation.Slides.Count+1,11);
                else
                    %ppt.currentSlide = ppt.presentation.Slides.AddSlide(slideIndex,...
                    %    ppt.presentation.SlideMaster.CustomLayouts.Item(6));
                    ppt.currentSlide = invoke(ppt.presentation.Slides,'Add',...
                        slideIndex,11);
                end
                %select the first slide to allow setting of the slide
                ppt.currentSlide.Select;
            catch
                error('adding slide failed')
            end
        end
        
        function selectSlide(ppt,slideIndex)
            %% selectSlide(slideIndex)
            % adds a new slide to the power point presentation
            % if slideIndex is empty, will add to the end of the
            % presentation, if slideIndex is not will add it there
            
            try
                ppt.currentSlide = ppt.presentation.Slides.Item(slideIndex);
                %select the first slide to allow setting of the slide
                ppt.currentSlide.Select;
            catch
                error('Selecting slide failed')
            end
        end
        
        function save(ppt,savePath,varargin)
            %% save(savePath)
            % saves current power point presentation to the path
            % provided by savePath, can be an abosulte path or a relative
            % path
            %
            % Allowable additional variables to pass in
            % Password - 'password' can be used to add a password to the
            %   power point presentation when saving
            
            password = '';
            for a = 1:2:nargin-2
                %case statement to parse out values
                switch upper(varargin{a})
                    case 'PASSWORD'
                        password = varargin{a+1};
                        if(~ischar(password))
                            error('Password must be a character array')
                        end
                    otherwise
                        warning('Invalid input detected')
                end
            end
            try
                %check ending of path provided for power point extention,
                %add if doesn't exist
                if(strcmpi(savePath(end-5:end),'.pptx'))
                    ending = '';
                else
                    ending = '.pptx';
                end
                %if no password passed, save without password
                if(~isempty(password))
                    ppt.presentation.Password = password;
                end
                ppt.presentation.SaveAs([savePath ending]);
            catch
                error('Saving failed, check path provided')
            end
            
        end
        
        function close(ppt)
            %% close
            % closes the presentation
            ppt.presentation.Close;
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
        
        function addFormattedText(ppt, textRange, text, suppres_newline)
            %% interprets the text and the new text to the slide
            
            %seperates out the special text from the rest
            %[special,start_loc,end_loc] = regexp(text,'<s ([a-zA-Z-0-9]+:[a-zA-Z- 0-9]+[;]*)*>','tokens');
            %finds the end to the text format
            [tag_type,tag_start,tag_end] = regexp(text,'<([\\//]*[buis])([a-zA-Z -;:0-9]*)>','tokens');
            
            if(nargin < 4)
                suppres_newline = 0;
            end
            %[term_start_loc,term_end_loc] = regexp(text,'<[\\//]s>');
            %[bui_b,bui_b_start,bui_b_finish] = regexp(text,'(<[\\//][bui]>)','tokens');
            %[bui_e,bui_e_start,bui_e_finish] = regexp(text,'(<[bui]>)','tokens');
            
            %start the struct for formatting the text
            text_format(1).font = 'Times New Roman';
            text_format(1).color = 'black';
            text_format(1).bg = 'white';
            text_format(1).size = 20;
            text_format(1).href = '';
            text_format(1).bold = 0;
            text_format(1).underlined = 0;
            text_format(1).italicize = 0;
            text_format(1).text = '';
            
            index_text = 1;
            index_tag = 1;
            referece_text = 1;
            
            % handle text before any tags
            if(~isempty(tag_start) && tag_start(1) > 1)
                text_format(1).text = text(1:tag_start(1));
            end
            
            while(index_tag <= length(tag_start))
    
                switch lower(tag_type{index_tag}{1})
                    case 's'
                        index_text = index_text+1;
                        text_format(index_text) = text_format(referece_text);
                        special_broken = regexp(tag_type{index_tag}{2},'([a-zA-Z-0-9]+):([a-zA-Z- 0-9]+)[;]*','tokens');
                        for b = 1:length(special_broken)
                            switch lower(special_broken{b}{1})
                                case 'font-family'
                                    text_format(index_text).font = special_broken{b}{2};
                                case 'font-size'
                                    text_format(index_text).size = str2num(special_broken{b}{2});
                                case 'bg'
                                    text_format(index_text).bg = special_broken{b}{2};
                                case 'href'
                                    text_format(index_text).href = special_broken{b}{2};
                                case 'color'
                                    text_format(index_text).color = special_broken{b}{2};
                                otherwise
                                    warning('A not covered case was introduced in the special formatting');
                            end
                        end
                        referece_text = referece_text+1;
                    case 'b'
                        index_text = index_text+1;
                        text_format(index_text) = text_format(referece_text);
                        text_format(index_text).bold = 1;
                        referece_text = referece_text+1;
                    case 'i'
                        index_text = index_text+1;
                        text_format(index_text) = text_format(referece_text);
                        text_format(index_text).italicize = 1;
                        referece_text = referece_text+1;
                    case 'u'
                        index_text = index_text+1;
                        text_format(index_text) = text_format(referece_text);
                        text_format(index_text).underlined = 1;
                        referece_text = referece_text+1;
                    case {'/u' '\u' '/s' '\s' '/b' '\b' '/i' '\i'}
                        index_text = index_text+1;
                        referece_text = referece_text-1;
                        text_format(index_text) = text_format(referece_text);
                    otherwise
                end
                
                index_tag = index_tag+1;
                if(index_tag < length(tag_start) && ...
                        (tag_start(index_tag)+1 == tag_start(index_tag-1)))
                    index_text = index_text-1;
                end
                if(index_tag <= length(tag_start))
                    text_format(index_text).text = text(tag_end(index_tag-1)+1:tag_start(index_tag)-1);
                end
            end
            
            %grab the last text if its not within a tag
            if(~isempty(tag_end) && (tag_end(end)+1 <= length(text)))
                text_format(index_text).text = text(tag_end(end)+1:length(text));
            elseif(isempty(tag_start) && isempty(tag_end))
                text_format(index_text).text = text;
            else
                text_format(index_text).text = '';
            end
            
            %desired to not add newlines to table, but to add them when
            %making multiple lines for text
            if(~suppres_newline)
                %add on a newline character to the last non-empty text box
                for a = 0:length(text_format)-1
                    if(~isempty(text_format(index_text-a).text) || a == length(text_format)-1)
                        text_format(index_text-a).text = [text_format(index_text-a).text char(13)];
                        break;
                    end
                end
            end
            
            %add text format to the slide
            for a = 1:length(text_format)
                if(~isempty(text_format(a).text))
                    start_text = length(textRange.Text)+1;
                    end_text = start_text + length(text_format(a).text);
                    textRange.InsertAfter(text_format(a).text);
                    try
                        %specify the font name
                        textRange.Characters(start_text,end_text).Font.Name = text_format(a).font;
                        %specify the font size
                        textRange.Characters(start_text,end_text).Font.Size = text_format(a).size;
                        %specify the font color
                        textRange.Characters(start_text,end_text).Font.Color.RGB = rgb(text_format(a).color,'ppt');
                        %make characters underlined
                        if(text_format(a).underlined)
                            textRange.Characters(start_text,end_text).Font.Underline = 'msoTrue';
                        else
                            textRange.Characters(start_text,end_text).Font.Underline = 'msoFalse';
                        end
                        %make characters bold
                        if(text_format(a).bold)
                            textRange.Characters(start_text,end_text).Font.Bold = 'msoTrue';
                        else
                            textRange.Characters(start_text,end_text).Font.Bold = 'msoFalse';
                        end
                        %make characters italic
                        if(text_format(a).italicize)
                            textRange.Characters(start_text,end_text).Font.Italic = 'msoTrue';
                        else
                            textRange.Characters(start_text,end_text).Font.Italic = 'msoFalse';
                        end
                        %Add the href link
                        if(~isempty(text_format(a).href))
                            textRange.Characters(start_text,end_text).ActionSettings.Item(1).Hyperlink.SubAddress = ...
                                text_format(a).href;
                        end
                        
                    catch
                        warning('Failure to addtext')
                    end
                end
            end
            
        end
    end
    
end