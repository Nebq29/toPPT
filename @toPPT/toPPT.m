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
                    ppt.currentSlide.SlideNumber,SectionTitle);
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
        
        function removeSlide(ppt,slideIndex)
            try
                remslides = fliplr(sort(slideIndex));
                for a = 1:length(slideIndex)
                    ppt.presentation.Slides.Item(remslides(a)).Delete
                end
            catch
                error('error removing slides')
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
        
        function [titles] = getSlideTitles(ppt)
            for a = 1:ppt.presentation.Slides.Count
                titles{a} = ppt.presentation.Slides.Item(a).Shapes.Title.TextFrame.TextRange.Text;
            end
        end
        
        function [sections, slides] = getSlideSections(ppt)
            slides = [];
            sections = {};
            for a = 1:ppt.presentation.SectionProperties.Count
                sections{a} = ppt.presentation.SectionProperties.Name(a);
                slides(a) = ppt.presentation.SectionProperties.SlidesCount(a);
            end
        end
        
        function addFormattedText(ppt, textRange, text, suppres_newline)
            %% interprets the text and the new text to the slide
            
            %seperates out the special text from the rest
            %[special,start_loc,end_loc] = regexp(text,'<s ([a-zA-Z-0-9]+:[a-zA-Z- 0-9]+[;]*)*>','tokens');
            %finds the end to the text format
            text = strrep(text,'–','-'); %catch the auto replaced - in ppt
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
                        special_broken = regexp(tag_type{index_tag}{2},'([a-zA-Z-0-9]+):([",@a-zA-Z- 0-9]+)[;]*','tokens');
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
                        textRange.Characters(start_text,end_text).Font.Color.RGB = ppt.rgb(text_format(a).color,'ppt');
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
        
        function lines = findPPTNumLines(ppt,text,width)
            %returns the expected height of the box in mm
            maxcharperline = width * 1.8;
            newlinechar = [1 regexp(text,char(10))];
            if(length(newlinechar)<2)
                lines = 4*sum(ceil(length(text)/maxcharperline))+2;
            else
                lines = 4*sum(ceil(diff(newlinechar)/maxcharperline))+6;
            end
        end

        function rgbOut = rgb(ppt,s,outputFormat)
            
            %% Possible inputs for s is a colorName or a hexValue like '#F0FFF0'
            
            isHex = 0;
            % First check if colorname or hexvalue
            if(strcmp(s(1),'#'))
                isHex = 1;
                % Delete #
                s = s(2:end);
                % Translate to Hex structure
                rgb{1} = s(1:2);
                rgb{2} = s(3:4);
                rgb{3} = s(5:6);
            else
                % We just have to look up the hexcolor in the table
                css = {
                    %White colors
                    'FF','FF','FF', 'White'
                    'FF','FA','FA', 'Snow'
                    'F0','FF','F0', 'Honeydew'
                    'F5','FF','FA', 'MintCream'
                    'F0','FF','FF', 'Azure'
                    'F0','F8','FF', 'AliceBlue'
                    'F8','F8','FF', 'GhostWhite'
                    'F5','F5','F5', 'WhiteSmoke'
                    'FF','F5','EE', 'Seashell'
                    'F5','F5','DC', 'Beige'
                    'FD','F5','E6', 'OldLace'
                    'FF','FA','F0', 'FloralWhite'
                    'FF','FF','F0', 'Ivory'
                    'FA','EB','D7', 'AntiqueWhite'
                    'FA','F0','E6', 'Linen'
                    'FF','F0','F5', 'LavenderBlush'
                    'FF','E4','E1', 'MistyRose'
                    %Grey colors'
                    '80','80','80', 'Gray'
                    'DC','DC','DC', 'Gainsboro'
                    'D3','D3','D3', 'LightGray'
                    'C0','C0','C0', 'Silver'
                    'A9','A9','A9', 'DarkGray'
                    '69','69','69', 'DimGray'
                    '77','88','99', 'LightSlateGray'
                    '70','80','90', 'SlateGray'
                    '2F','4F','4F', 'DarkSlateGray'
                    '00','00','00', 'Black'
                    %Red colors
                    'FF','00','00', 'Red'
                    'FF','A0','7A', 'LightSalmon'
                    'FA','80','72', 'Salmon'
                    'E9','96','7A', 'DarkSalmon'
                    'F0','80','80', 'LightCoral'
                    'CD','5C','5C', 'IndianRed'
                    'DC','14','3C', 'Crimson'
                    'B2','22','22', 'FireBrick'
                    '8B','00','00', 'DarkRed'
                    %Pink colors
                    'FF','C0','CB', 'Pink'
                    'FF','B6','C1', 'LightPink'
                    'FF','69','B4', 'HotPink'
                    'FF','14','93', 'DeepPink'
                    'DB','70','93', 'PaleVioletRed'
                    'C7','15','85', 'MediumVioletRed'
                    %Orange colors
                    'FF','A5','00', 'Orange'
                    'FF','8C','00', 'DarkOrange'
                    'FF','7F','50', 'Coral'
                    'FF','63','47', 'Tomato'
                    'FF','45','00', 'OrangeRed'
                    %Yellow colors
                    'FF','FF','00', 'Yellow'
                    'FF','FF','E0', 'LightYellow'
                    'FF','FA','CD', 'LemonChiffon'
                    'FA','FA','D2', 'LightGoldenrodYellow'
                    'FF','EF','D5', 'PapayaWhip'
                    'FF','E4','B5', 'Moccasin'
                    'FF','DA','B9', 'PeachPuff'
                    'EE','E8','AA', 'PaleGoldenrod'
                    'F0','E6','8C', 'Khaki'
                    'BD','B7','6B', 'DarkKhaki'
                    'FF','D7','00', 'Gold'
                    %Brown colors
                    'A5','2A','2A', 'Brown'
                    'FF','F8','DC', 'Cornsilk'
                    'FF','EB','CD', 'BlanchedAlmond'
                    'FF','E4','C4', 'Bisque'
                    'FF','DE','AD', 'NavajoWhite'
                    'F5','DE','B3', 'Wheat'
                    'DE','B8','87', 'BurlyWood'
                    'D2','B4','8C', 'Tan'
                    'BC','8F','8F', 'RosyBrown'
                    'F4','A4','60', 'SandyBrown'
                    'DA','A5','20', 'Goldenrod'
                    'B8','86','0B', 'DarkGoldenrod'
                    'CD','85','3F', 'Peru'
                    'D2','69','1E', 'Chocolate'
                    '8B','45','13', 'SaddleBrown'
                    'A0','52','2D', 'Sienna'
                    '80','00','00', 'Maroon'
                    %Green colors
                    '00','80','00', 'Green'
                    '98','FB','98', 'PaleGreen'
                    '90','EE','90', 'LightGreen'
                    '9A','CD','32', 'YellowGreen'
                    'AD','FF','2F', 'GreenYellow'
                    '7F','FF','00', 'Chartreuse'
                    '7C','FC','00', 'LawnGreen'
                    '00','FF','00', 'Lime'
                    '32','CD','32', 'LimeGreen'
                    '00','FA','9A', 'MediumSpringGreen'
                    '00','FF','7F', 'SpringGreen'
                    '66','CD','AA', 'MediumAquamarine'
                    '7F','FF','D4', 'Aquamarine'
                    '20','B2','AA', 'LightSeaGreen'
                    '3C','B3','71', 'MediumSeaGreen'
                    '2E','8B','57', 'SeaGreen'
                    '8F','BC','8F', 'DarkSeaGreen'
                    '22','8B','22', 'ForestGreen'
                    '00','64','00', 'DarkGreen'
                    '6B','8E','23', 'OliveDrab'
                    '80','80','00', 'Olive'
                    '55','6B','2F', 'DarkOliveGreen'
                    '00','80','80', 'Teal'
                    %Blue colors
                    '00','00','FF', 'Blue'
                    'AD','D8','E6', 'LightBlue'
                    'B0','E0','E6', 'PowderBlue'
                    'AF','EE','EE', 'PaleTurquoise'
                    '40','E0','D0', 'Turquoise'
                    '48','D1','CC', 'MediumTurquoise'
                    '00','CE','D1', 'DarkTurquoise'
                    'E0','FF','FF', 'LightCyan'
                    '00','FF','FF', 'Cyan'
                    '00','FF','FF', 'Aqua'
                    '00','8B','8B', 'DarkCyan'
                    '5F','9E','A0', 'CadetBlue'
                    'B0','C4','DE', 'LightSteelBlue'
                    '46','82','B4', 'SteelBlue'
                    '87','CE','FA', 'LightSkyBlue'
                    '87','CE','EB', 'SkyBlue'
                    '00','BF','FF', 'DeepSkyBlue'
                    '1E','90','FF', 'DodgerBlue'
                    '64','95','ED', 'CornflowerBlue'
                    '41','69','E1', 'RoyalBlue'
                    '00','00','CD', 'MediumBlue'
                    '00','00','8B', 'DarkBlue'
                    '00','00','80', 'Navy'
                    '19','19','70', 'MidnightBlue'
                    %Purple colors
                    '80','00','80', 'Purple'
                    'E6','E6','FA', 'Lavender'
                    'D8','BF','D8', 'Thistle'
                    'DD','A0','DD', 'Plum'
                    'EE','82','EE', 'Violet'
                    'DA','70','D6', 'Orchid'
                    'FF','00','FF', 'Fuchsia'
                    'FF','00','FF', 'Magenta'
                    'BA','55','D3', 'MediumOrchid'
                    '93','70','DB', 'MediumPurple'
                    '99','66','CC', 'Amethyst'
                    '8A','2B','E2', 'BlueViolet'
                    '94','00','D3', 'DarkViolet'
                    '99','32','CC', 'DarkOrchid'
                    '8B','00','8B', 'DarkMagenta'
                    '6A','5A','CD', 'SlateBlue'
                    '48','3D','8B', 'DarkSlateBlue'
                    '7B','68','EE', 'MediumSlateBlue'
                    '4B','00','82', 'Indigo'
                    %Gray repeated with spelling grey
                    '80','80','80', 'Grey'
                    'D3','D3','D3', 'LightGrey'
                    'A9','A9','A9', 'DarkGrey'
                    '69','69','69', 'DimGrey'
                    '77','88','99', 'LightSlateGrey'
                    '70','80','90', 'SlateGrey'
                    '2F','4F','4F', 'DarkSlateGrey'
                    };
                num = css(:,1:3);
                name = css(:,4);
                k = find(strcmpi(s, name));
                if isempty(k)
                    error(['Unknown color: ' s]);
                else
                    rgb = num(k(1), :);
                end
            end
            
            %% Now we can proceed - depening on the outpurformat we need
            switch outputFormat
                case 'dec'
                    
                    %         num = reshape(hex2dec(num), [], 3);
                    %         % Divide most numbers by 256 for "aesthetic" reasons (green=[0 0.5 0])
                    %         I = num < 240;  % (interpolate F0--FF linearly from 240/256 to 1.0)
                    %         num(I) = num(I)/256;
                    %         num(~I) = ((num(~I) - 240)/15 + 15)/16; + 240;
                    %
                    %         num = num*256;
                    %         rgbCell = num(k(1), :);
                    rgbOut = hex2dec([rgb{1},rgb{2},rgb{3}]);
                    
                case 'hex'
                    rgbOut = [rgb{1},rgb{2},rgb{3}];
                case 'ppt'
                    % For some strange reason ppt is desiged the following way BRG
                    % instead of RGB that means that FF0000 is transfered into 00FF00
                    % and % instead of RGB that means that 0000FF is transfered into FF0000
                    % so basically it is flipped
                    rgbOut = hex2dec([rgb{3},rgb{2},rgb{1}]);
                    
                case 'decVec'
                    
                    r1 = hex2dec(rgb{1});
                    r2 = hex2dec(rgb{2});
                    r3 = hex2dec(rgb{3});
                    
                    rgbOut = [r1,r2,r3];
                    
            end
            
        end

    end
    
end