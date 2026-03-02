function [refName, opts] = reportOptionsGUI(channelNames, opts)
% psd.ui.reportOptionsGUI
% Un único diálogo para:
%  - seleccionar canal de referencia
%  - layout de paquetes (RxC)
%  - formato de imagen (svg/png; default svg)
%  - figura inicial y página inicial
%  - plantilla Word (opcional)

if nargin < 2 || isempty(opts)
    opts = struct();
end
opts = psd.local.applyDefaults(opts);

if isempty(channelNames)
    error('reportOptionsGUI:EmptyChannels','No hay canales para seleccionar.');
end

% -------------------------
% UI (uifigure modal)
% -------------------------
f = uifigure('Name','Opciones de reporte', 'Position',[200 200 560 340], ...
    'WindowStyle','modal');

g = uigridlayout(f, [8 4]);
g.RowHeight = {24,24,24,24,24,24,24,'1x'};
g.ColumnWidth = {150,'1x',120,80};
g.Padding = [12 12 12 12];
g.RowSpacing = 8;
g.ColumnSpacing = 10;

% Canal de referencia
uilabel(g,'Text','Canal de referencia:','HorizontalAlignment','left');
ddRef = uidropdown(g, 'Items', channelNames, 'Value', channelNames{1});
ddRef.Layout.Column = [2 4];

% Layout de paquete
uilabel(g,'Text','Paquete (filas x cols):');
ddPack = uidropdown(g, 'Items', {'1x1','1x2','2x1','2x2'}, 'Value', sprintf('%dx%d',opts.PackageRows,opts.PackageCols));
ddPack.Layout.Column = 2;
uilabel(g,'Text','');
uilabel(g,'Text','');

% Formato
uilabel(g,'Text','Formato imágenes:');
ddFmt = uidropdown(g, 'Items', {'svg','png'}, 'Value', lower(opts.ImageFormat));
ddFmt.Layout.Column = 2;
uilabel(g,'Text','(default svg)','HorizontalAlignment','left');
uilabel(g,'Text','');

% Figura inicial
uilabel(g,'Text','Figura inicial:');
efFig = uieditfield(g,'numeric','Value',opts.StartFigure,'Limits',[0 inf],'RoundFractionalValues','on');
efFig.Layout.Column = 2;
uilabel(g,'Text','');
uilabel(g,'Text','');

% Página inicial
uilabel(g,'Text','Página inicial:');
efPage = uieditfield(g,'numeric','Value',opts.StartPage,'Limits',[1 inf],'RoundFractionalValues','on');
efPage.Layout.Column = 2;
uilabel(g,'Text','');
uilabel(g,'Text','');

% Plantilla Word
uilabel(g,'Text','Plantilla Word (.docx):');
efTpl = uieditfield(g,'text','Value',opts.TemplatePath);
efTpl.Layout.Column = [2 3];
btnBrowse = uibutton(g,'Text','Buscar...');

btnBrowse.ButtonPushedFcn = @(~,~)localBrowseTemplate();
    function localBrowseTemplate()
        [file, path] = uigetfile({'*.docx','Word (*.docx)'}, 'Selecciona tu Word base');
        if isequal(file,0); return; end
        efTpl.Value = fullfile(path,file);
    end

% Nota tipografía fija
lbl = uilabel(g,'Text','Matlab: Times New Roman, 13 pt (fijo).','FontAngle','italic');
lbl.Layout.Row = 7;
lbl.Layout.Column = [1 4];

% Botones
pBtns = uipanel(g); pBtns.Layout.Row = 8; pBtns.Layout.Column = [1 4];
gb = uigridlayout(pBtns,[1 3]);
gb.ColumnWidth = {'1x',100,100};
gb.Padding = [0 0 0 0];

uilabel(gb,'Text','');
btnCancel = uibutton(gb,'Text','Cancelar');
btnOk     = uibutton(gb,'Text','OK','ButtonPushedFcn',@(~,~)uiresume(f));

btnCancel.ButtonPushedFcn = @(~,~)localCancel();
    function localCancel()
        refName = '';
        opts = struct();
        delete(f);
    end

uiwait(f);
if ~isvalid(f)
    % cancelado
    if isempty(refName)
        error('reportOptionsGUI:Cancelled','Operación cancelada por el usuario.');
    end
    return;
end

% -------------------------
% Return values
% -------------------------
refName = ddRef.Value;

pack = ddPack.Value;
tok = regexp(pack,'^(\d+)x(\d+)$','tokens','once');
opts.PackageRows = str2double(tok{1});
opts.PackageCols = str2double(tok{2});

opts.ImageFormat = lower(ddFmt.Value);

opts.FontName = 'Times New Roman';
opts.FontSize = 13;

opts.StartFigure = efFig.Value;
opts.StartPage   = efPage.Value;
opts.TemplatePath = strtrim(efTpl.Value);

delete(f);

end
