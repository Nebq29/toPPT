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
            %Opens an existing PPT
            try
                ppt.activeXCom.Presentations.Open(file);
                ppt.presentation = wordClass.activeXCom.ActivePresentation;
            catch
                error('Seems the file does not exist')
            end
        end
        
        function NewTemplate(ppt)
            %Opens the existing template
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
            %Creates a new blank document
            try
                ppt.presentation = ppt.activeXCom.Presentation.Add;
            catch
                error('unable to open new document')
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
    end
    
end