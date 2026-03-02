function imgPaths = plotAllChannels(channels, varargin)
% psd.plot.plotAllChannels
% Compatible con:
%   imgPaths = psd.plot.plotAllChannels(channels, baseFolder, opts)
%   imgPaths = psd.plot.plotAllChannels(channels, opts)

% -------------------------
% Parse de argumentos
% -------------------------
baseFolder = '';
opts = struct();

if numel(varargin) == 1
    opts = varargin{1};
    if isfield(opts,'BaseFolder')
        baseFolder = opts.BaseFolder;
    else
        baseFolder = pwd;
    end
elseif numel(varargin) == 2
    baseFolder = varargin{1};
    opts = varargin{2};
else
    error('plotAllChannels:InvalidInputs', ...
        'Use plotAllChannels(channels, opts) or plotAllChannels(channels, baseFolder, opts).');
end

% Defaults
if ~isfield(opts,'FontName') || isempty(opts.FontName)
    opts.FontName = 'Times New Roman';
end
if ~isfield(opts,'FontSize') || isempty(opts.FontSize)
    opts.FontSize = 13;
end
if ~isfield(opts,'ImageFormat') || isempty(opts.ImageFormat)
    opts.ImageFormat = 'svg';
end

% En Matlab: todo a 13 pt (título, ejes, labels)
opts.FontSizeTitle = opts.FontSize;
opts.FontSizeLabel = opts.FontSize;
opts.FontSizeAxes  = opts.FontSize;

% -------------------------
% Carpeta de export
% -------------------------
exportFolder = fullfile(baseFolder, 'FIGURES');
if ~exist(exportFolder,'dir'); mkdir(exportFolder); end

N = numel(channels);

% imgPaths: rutas que se usan para armar el Word.
% Si el usuario pide SVG, igual exportamos un PNG "compañero" (600 dpi)
% para poder recortar e insertar en Word.
imgPaths = strings(N,1);

lineColor = [0 0.4471 0.7412];

for i = 1:N

    % Número de canal (si el nombre termina en dígitos, úsalo; si no, usa i)
    chName = channels(i).name;
    chNumStr = regexp(chName,'\d+$','match','once');
    if isempty(chNumStr)
        chNum = i;
        chNumStr = num2str(i);
    else
        chNum = str2double(chNumStr);
        if isnan(chNum); chNum = i; chNumStr = num2str(i); end
    end

    unit = localUnitFromChannel(chNum);

    baseName = sprintf('Channel_%02d', chNum);
    if strcmpi(opts.ImageFormat,'png')
        imgFile = fullfile(exportFolder, baseName + ".png");
        imgPaths(i) = string(imgFile);
        svgFile = ""; %#ok<NASGU>
        pngFile = ""; %#ok<NASGU>
    else
        % Guardar SVG por defecto + PNG para Word
        svgFile = fullfile(exportFolder, baseName + ".svg");
        pngFile = fullfile(exportFolder, baseName + ".png");
        imgPaths(i) = string(pngFile);
    end

    % Figura temporal grande (evita deformación)
    f = figure('Visible','off','Color','w','Units','pixels','Position',[100 100 1800 950]);
    ax = axes('Parent', f);

    t = channels(i).time;
    y = channels(i).signal;

    plot(ax, t, y, 'Color', lineColor, 'LineWidth', 1.2);
    grid(ax, 'on');

    title(ax, sprintf('Channel %d', chNum), ...
        'FontName', opts.FontName, 'FontSize', opts.FontSizeTitle);

    xlabel(ax, 't (sec)', ...
        'FontName', opts.FontName, 'FontSize', opts.FontSizeLabel);

    % ylabel con subíndice + unidad
    if strcmp(unit,'-')
        ylab = sprintf('signal_{%d}', chNum);
    else
        ylab = sprintf('signal_{%d} (%s)', chNum, unit);
    end

    ylabel(ax, ylab, ...
        'FontName', opts.FontName, 'FontSize', opts.FontSizeLabel, ...
        'Interpreter','tex');

    set(ax, ...
        'FontName', opts.FontName, ...
        'FontSize', opts.FontSizeAxes, ...
        'Box','off');

    % Maximizar contenido y reducir borde blanco
    set(ax, 'Units','normalized', 'Position',[0.08 0.12 0.90 0.82]);
    drawnow;
    try
        set(ax, 'LooseInset', get(ax,'TightInset'));
    catch
    end
    drawnow;

    % Export estable
    set(f, 'PaperPositionMode', 'auto');
    drawnow;

    if strcmpi(opts.ImageFormat,'png')
        print(f, char(imgPaths(i)), '-dpng', '-r600');
    else
        try
        
% Tight layout to minimize margins and keep ~1 mm equivalent border
set(gca,'LooseInset',max(get(gca,'TightInset'),0.01*[1 1 1 1]));
set(gcf,'PaperPositionMode','auto');
drawnow;
print(gcf, char(svgFile), '-dsvg');
    % Also export EMF for Word (vector)
try
    emfFile = replace(char(svgFile), '.svg', '.emf');
    print(gcf, emfFile, '-dmeta');
catch
end
catch ME
        % Fallback for MATLAB versions where exportgraphics does not support SVG.
        % Use PRINT on the parent figure.
        fig = ancestor(ax,'figure');
        print(fig, char(svgFile), '-dsvg');
    end
        print(f, char(pngFile), '-dpng', '-r600');
    end

    close(f);
end

end

% -------- helper local --------
function unit = localUnitFromChannel(chNum)
if chNum == 1
    unit = 'kN';
elseif chNum == 2
    unit = 'mm';
elseif chNum == 3
    unit = 'kN';
elseif chNum == 4 || chNum == 5
    unit = '-';
elseif chNum >= 6 && chNum <= 16
    unit = 'g';
else
    unit = '-';
end
end
