function opts = applyDefaults(opts)
% psd.local.applyDefaults
% Centraliza defaults para las opciones del reporte/figuras.

if nargin < 1 || isempty(opts)
    opts = struct();
end

% ---- defaults de control ----
if ~isfield(opts,'ExportWord') || isempty(opts.ExportWord)
    opts.ExportWord = true;
end

% ---- layout de paquetes ----
if ~isfield(opts,'PackageRows') || isempty(opts.PackageRows) || ~isnumeric(opts.PackageRows)
    opts.PackageRows = 2;
end
if ~isfield(opts,'PackageCols') || isempty(opts.PackageCols) || ~isnumeric(opts.PackageCols)
    opts.PackageCols = 2;
end
opts.PackageRows = max(1, round(opts.PackageRows));
opts.PackageCols = max(1, round(opts.PackageCols));

% ---- export de imágenes ----
if ~isfield(opts,'ImageFormat') || isempty(opts.ImageFormat)
    opts.ImageFormat = 'svg';
end
opts.ImageFormat = lower(string(opts.ImageFormat));
if opts.ImageFormat ~= "svg" && opts.ImageFormat ~= "png"
    opts.ImageFormat = "svg";
end
opts.ImageFormat = char(opts.ImageFormat);

% ---- tipografía Matlab ----
if ~isfield(opts,'FontName') || isempty(opts.FontName)
    opts.FontName = 'Times New Roman';
end
if ~isfield(opts,'FontSize') || isempty(opts.FontSize) || ~isnumeric(opts.FontSize)
    opts.FontSize = 13;
end
opts.FontSize = max(1, round(opts.FontSize));

% ---- numeración ----
if ~isfield(opts,'StartFigure') || isempty(opts.StartFigure) || ~isnumeric(opts.StartFigure)
    opts.StartFigure = 1;
end
opts.StartFigure = max(0, round(opts.StartFigure));

if ~isfield(opts,'StartPage') || isempty(opts.StartPage) || ~isnumeric(opts.StartPage)
    opts.StartPage = 1;
end
opts.StartPage = max(1, round(opts.StartPage));

% ---- plantilla Word ----
if ~isfield(opts,'TemplatePath') || isempty(opts.TemplatePath)
    opts.TemplatePath = '';
end

end
