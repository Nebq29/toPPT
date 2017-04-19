function addImage(ppt, fig, varargin)
    %%addImage(fig, varargin)
    % fig is either the matlab figure to be added to the slide or the
    % location of an image to be added to the slide
    %
    % The default figure format is .png to help reduce overall size of the
    % power point, but other possible export formats are allowable as well
    % 
    % Allowable additional variables to pass in
    % Location - [%x,%y] where this is from the NW corner of the slide,
    %   defaults to [5,25]
    % Size - [x%,y%] gives the size of the image to be created defaults
    %   to [90, 70]
    % Format - 'png', 'jpeg', ...etc for a full list check <a href="matlab:helpview([docroot '/matlab/ref/print.html#inputarg_formattype'])">here</a>
    %   defaults to png
    
    %Valid formats and corresponding file extentions of formats
    valid_formats = {'jpeg', 'png', 'tiff', 'tiffn', 'meta', 'bmpmono', 'bmp',...
        'bmp16m', 'bmp256', 'hdf', 'pbm', 'pbmraw', 'pcxmono', 'pcx24b',...
        'pcx256', 'pcx16', 'pgm', 'pgmraw', 'ppm', 'ppmraw', 'eps', 'epsc',...
        'eps2', 'epsc2', 'meta', 'svg', 'ps', 'psc', 'ps2', 'psc2'};
    format_extentions = {'.jpg','.png','.tif','.tif','.emf','.bmp','.bmp',...
        '.bmp','.bmp','.hdf','.pbm','.pbm','.pcx','.pcx','.pcx','.pcx','.pgm',...
        '.pgm','.pgm','.pgm','.eps','.eps','.eps','.eps','.emf','.svg','.ps',...
        '.ps','.ps','.ps'};
    
    %set default values for function
    location = [5,25];
    boxSize = [90, 70];
    format = 'png';
    format_ext = '.png';
    for a = 1:2:nargin-2
        %case statement to parse out values
        switch upper(varargin{a})
            case 'LOCATION'
                location = varargin{a+1};
                if(numel(location) ~= 2)
                    error('Location must be a 2 value array')
                elseif(~isnumeric(location))
                    error('Location must be passed as a numeric value')
                elseif(sum(location > 100 | location < 0))
                    error('Location must be values between 0 and 100')
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
            case 'FORMAT'
                format = varargin{a+1};
                if(~sum(strcmpi(format,valid_formats)))
                    error('Invalid format passed in')
                end
                format_ext = format_extentions{strcmpi(format,valid_formats)};
            otherwise
                warning('Invalid input detected')
        end
    end
    slideHeight = ppt.presentation.PageSetup.SlideHeight;
    slideWidth = ppt.presentation.PageSetup.SlideWidth;
    
    %check validaty of input and convert to correct image format
    if(ishandle(fig))
        pathImage = fullfile(tempdir, ['pptfig',format_ext]);
        fig.PaperPositionMode = 'auto';
        print(fig,['-d' format],pathImage);
    elseif(ischar(fig) && exists(fig))
        pathImage = fig;
    else
        error('Passed figure must be a valid file or figure handle')
    end
    
    %Create the Imagebox
    img = ppt.presentation.Slides.Item(ppt.presentation.Slides.Count).Shapes.AddPicture(pathImage,'msoFalse','msoTrue',...
        slideWidth*location(1)/100,slideHeight*location(2)/100,slideWidth*boxSize(1)/100,slideHeight*boxSize(2)/100);
    
    %send the image to the back
    img.ZOrder('msoSendToBack');
        
end