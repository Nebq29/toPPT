function addTableOfContents(ppt,h)
% Generates a table of contents based on sections and titles of each slide.
% A slide with a (continued) will not generate a slide number, while a
% section will create a section header for the following slides.
%
% Passing a power point title in will open the specific power point
% presentation while not passing will use the current active power point
% opened

    if(nargin>1)
        waitbar(0,h,'Clearing Existing Table of Contents')
    end
    [titles] = ppt.getSlideTitles;
    [sections,~] = ppt.getSlideSections;

    if(sum(cell2mat(regexp(titles,'Table of Contents'))))
        %has an existing table of contents, remove it

        for a = 0:ppt.presentation.Slides.Count-1
            if(~isempty(regexp(titles{length(titles)-a},'Table of Contents','once')))
                ppt.presentation.Slides.Item(length(titles)-a).Delete;
            end
        end
        for a = 0:ppt.presentation.SectionProperties.Count-1
            if(~isempty(regexp(sections{length(sections)-a},'Table of Contents','once')))
                ppt.presentation.SectionProperties.Delete(length(sections)-a,false);
            end
        end
    end
    
    if(nargin>1)
        waitbar(0,h,'Creating New Table of Contents')
    end
    [sections,slides] = ppt.getSlideSections;
    [titles] = ppt.getSlideTitles;

    table = {};
    slidenumbers = [];
    b = 1;
    bulletindent = [];
    bullettype = [];
    for a = 2:(length(titles))
        if((sum(slides(1:b))+1) == a)
            if(a ~= 2)
                table = {table{:} ...
                    ['<s font-family:Times New Roman;font-size:22> <\s>']};
                bulletindent = [bulletindent,1];
                bullettype = [bullettype,0];
                slidenumbers = [slidenumbers -1];
            end
            table = {table{:} ...
                ['<s font-family:Times New Roman;font-size:22><b>' sections{b+1} '<\b><\s>']};
            slidenumbers = [slidenumbers -1];
            b = b+1;
            bulletindent = [bulletindent,1];
            bullettype = [bullettype,0];
        end
        if(isempty(regexp(titles{a},'continued','once')))
            table = {table{:} ...
                ['<s font-family:Times New Roman;font-size:18;shref:255,@,' titles{a}...
                '>' titles{a} '<\s>']};
            slidenumbers = [slidenumbers a];
            bulletindent = [bulletindent,2];
            bullettype = [bullettype,0];
        end
    end
    
    ppt.newSlide(2);
    ppt.addSection('Table of Contents')
    ppt.setTitle('Table of Contents')
    
    if(length(table)>21)
        b = ceil(length(table)/21);
    else
        b = 1;
    end
    slidenum = {};
    for a = 1:length(slidenumbers)
            if(slidenumbers(a)<0)
                slidenum = {slidenum{:}...
                    ['<s font-family:Times New Roman;font-size:22> <\s>']};
            else
                slidenum = {slidenum{:}...
                    ['<s font-family:Times New Roman;font-size:18>' num2str(slidenumbers(a)+b) '<\s>']};
                table{a} = strrep(table{a},'@',num2str(slidenumbers(a)+b));
            end
    end
    for a = 1:b
        if(nargin>1)
            waitbar(a/b,h);
        end
        if(a~=1)
            ppt.newSlide(1+a);
            ppt.setTitle('Table of Contents (continued)')
        end
        if(a ~= b)
            ppt.addText(table(21*(a-1)+1:21*(a)),'Indent',bulletindent(21*(a-1)+1:21*(a)),...
                'Bullets',bullettype(21*(a-1)+1:21*(a)),'Location',[5,25],'Size',[95,32]);
            ppt.addText(slidenum(21*(a-1)+1:21*(a)),...
                'Bullets',bullettype(21*(a-1)+1:21*(a)),'Location',[75,25],'Size',[10,32]);
        else
            ppt.addText(table(21*(a-1)+1:end),'Indent',bulletindent(21*(a-1)+1:end),...
                'Bullets',bullettype(21*(a-1)+1:end),'Location',[5,25],'Size',[95,32]);
            ppt.addText(slidenum(21*(a-1)+1:end),...
                'Bullets',bullettype(21*(a-1)+1:end),'Location',[75,25],'Size',[10,32]);
        end
    end
    
    if(nargin>1)
        waitbar(0,h,'Updating Page Numbers')
    end
    if(~isempty(find(titles{1} == char(10),1,'first')))
        part_title = titles{1}(1:find(titles{1} == char(10),1,'first')-1);
    else
        part_title = '';
    end
    
    for a = 2:ppt.presentation.Slides.Count
        if(nargin>1)
            waitbar(a/ppt.presentation.Slides.Count,h);
        end
        ppt.presentation.Slides.Item(a).HeadersFooters.DateAndTime.Visible = 'msoTrue';
        ppt.presentation.Slides.Item(a).HeadersFooters.DateAndTime.Text = date;
        ppt.presentation.Slides.Item(a).HeadersFooters.SlideNumber.Visible = 'msoTrue';
        ppt.presentation.Slides.Item(a).HeadersFooters.Footer.Visible = 'msoTrue';
        ppt.presentation.Slides.Item(a).HeadersFooters.Footer.Text = part_title;
    end
    

end