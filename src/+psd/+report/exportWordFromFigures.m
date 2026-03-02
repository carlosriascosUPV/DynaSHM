function reportPath = exportWordFromFigures(imgPaths, baseFolder, opts)
% psd.report.exportWordFromFigures
% Arma un Word con "paquetes" (tablas) RxC, letras a) b) ... bajo cada imagen,
% y caption al final del paquete con los nombres de las imágenes.
%
% Requisitos implementados:
%   - Paquete RxC (1x1, 1x2, 2x1, 2x2)
%   - Letras en cursiva y paréntesis normal
%   - Caption: "Figura N. a) Nombre 1  +  b) Nombre 2 ..."
%   - Salto de página después de cada paquete
%   - 6 pt antes y después de cada paquete; 6 pt después del caption
%   - Auto-crop del borde blanco dejando 1 mm de borde
%   - Admite plantilla Word vacía con formato específico


% Default vertical spacing (points) used before/after each package and after caption
spacePts = 6;
try
    if isstruct(opts) && isfield(opts,'SpacePts') && ~isempty(opts.SpacePts)
        spacePts = opts.SpacePts;
    end
catch
end

import mlreportgen.dom.*

if nargin < 3 || isempty(opts)
    opts = struct();
end
opts = psd.local.applyDefaults(opts);

exportFolder = fullfile(baseFolder, 'REPORT');
if ~exist(exportFolder,'dir'); mkdir(exportFolder); end

fname = fullfile(exportFolder, ['Report (', datestr(now, 'dd-mm-yyyy HH-MM'), ')']);
if isfield(opts,'TemplatePath') && ~isempty(opts.TemplatePath)
    doc = Document(fname, 'docx', opts.TemplatePath);
else
    doc = Document(fname, 'docx');
end

% -------------------------
% Layout and typography
% -------------------------
R = max(1, round(opts.PackageRows));
C = max(1, round(opts.PackageCols));
perPackage = R*C;

fontName = 'Times New Roman';
fontPt = 13;

% Medidas para encajar bien en página A4 con márgenes típicos.
pageWidth_cm = 21;
margin_cm = 2.54;
usable_cm = pageWidth_cm - 2*margin_cm;
imgWidth_cm = (usable_cm * 0.98) / C;

% -------------------------
% Helpers
% -------------------------
letters = 'abcdefghijklmnopqrstuvwxyz';

% Asegurar imgPaths como cellstr
if isstring(imgPaths); imgPaths = cellstr(imgPaths); end
if ischar(imgPaths); imgPaths = {imgPaths}; end

N = numel(imgPaths);
figNum = opts.StartFigure;

% Espacio antes del paquete
preSpace = Paragraph(' ');
preSpace.Style = {OuterMargin('6pt','0pt','0pt','0pt')};

% -------------------------
% Main loop (packages)
% -------------------------
pkgStart = 1;
while pkgStart <= N

    pkgEnd = min(N, pkgStart + perPackage - 1);
    idx = pkgStart:pkgEnd;

    append(doc, localSpacer(spacePts));% Tabla RxC
    t = Table();
    t.Style = {Width('100%')};

    k = 0;
    for r = 1:R
        row = TableRow();
        for c = 1:C
            k = k + 1;
            cell = TableEntry();

            if k <= numel(idx)
                imgPath = imgPaths{idx(k)};

                % (1) Auto-crop (PNG/JPG/BMP/TIF). Si no se puede leer, se inserta tal cual.
                [imgForWord, wpx, hpx] = localCropKeeping1mm(imgPath);

                img = Image(imgForWord);
                img.Width = sprintf('%.2fcm', imgWidth_cm);
                img.Height = sprintf('%.2fcm', (hpx / wpx) * imgWidth_cm);

                pImg = Paragraph();
                pImg.Style = {HAlign('center')};
                append(pImg, img);

                % (2) Etiqueta a) b) ...
                pLbl = Paragraph();
                pLbl.Style = {HAlign('center'), FontFamily(fontName), FontSize(sprintf('%dpt',fontPt))};

                letter = letters(idx(k)-pkgStart+1);
                txtL = Text(letter); txtL.Italic = true;
                append(pLbl, txtL);
                append(pLbl, Text(')'));

                append(cell, pImg);
                append(cell, pLbl);
            else
                % Celda vacía
                append(cell, Paragraph(' '));
            end

            append(row, cell);
        end
        append(t, row);
    end

    append(doc, t);

    % Caption del paquete
    cap = Paragraph();
    cap.Style = {HAlign('center'), FontFamily(fontName), FontSize(sprintf('%dpt',fontPt)), OuterMargin('0pt','0pt','6pt','0pt')};

    append(cap, Text(sprintf('Figura %d. ', figNum)));

    for j = 1:numel(idx)
        name = localBaseName(imgPaths{idx(j)});
        letter = letters(j);

        txtL = Text(letter); txtL.Italic = true;
        append(cap, txtL);
        append(cap, Text(') '));
        append(cap, Text(name));

        if j < numel(idx)
            append(cap, Text('  +  '));
        end
    end

    append(doc, cap);

    % Espacio después del paquete (además del caption)
    postSpace = Paragraph(' ');
    postSpace.Style = {OuterMargin('0pt','0pt','6pt','0pt')};
    append(doc, localSpacer(spacePts));% Salto de página después de cada paquete
    append(doc, PageBreak());

    figNum = figNum + 1;
    pkgStart = pkgEnd + 1;
end

close(doc);
reportPath = [fname '.docx'];

end

% ============================================================
% Local helpers
% ============================================================
function name = localBaseName(p)
[~, name, ~] = fileparts(p);
end

function [tmpPng, wpx, hpx] = localCropKeeping1mm(imgPath)
% Recorta borde blanco dejando 1 mm alrededor.
% Devuelve una ruta PNG temporal que se puede insertar en Word.

tmpPng = imgPath;
wpx = 1; hpx = 1;

try
    I = imread(imgPath);
catch
    % No se puede leer (p.ej. SVG). Se inserta tal cual.
    try
        info = imfinfo(imgPath);
        wpx = info.Width; hpx = info.Height;
    catch
        wpx = 1000; hpx = 600;
    end
    return;
end

info = imfinfo(imgPath);
wpx = info.Width;
hpx = info.Height;

if size(I,3) == 3
    G = rgb2gray(I);
else
    G = I;
end

% máscara: no-blanco
mask = G < 250;

if ~any(mask(:))
    % todo blanco, dejar tal cual
    return;
end

[rows, cols] = find(mask);
top = min(rows); bottom = max(rows);
left = min(cols); right = max(cols);

% 1 mm de borde (en píxeles)
dpi = 600;
if isfield(info,'XResolution') && ~isempty(info.XResolution) && info.XResolution > 0
    dpi = info.XResolution;
end
borderPx = max(1, round(dpi / 25.4));

top = max(1, top - borderPx);
left = max(1, left - borderPx);
bottom = min(size(I,1), bottom + borderPx);
right  = min(size(I,2), right  + borderPx);

Icrop = I(top:bottom, left:right, :);

tmpPng = fullfile(tempdir, sprintf('crop_%s.png', char(java.util.UUID.randomUUID)));
imwrite(Icrop, tmpPng);

wpx = size(Icrop,2);
hpx = size(Icrop,1);

    
end

    
function p = localSpacer(nPts)
% Create a fresh empty paragraph with a given spacing-after (in points).
p = mlreportgen.dom.Paragraph();
try
    p.Style = {mlreportgen.dom.OuterMargin('0pt','0pt','0pt','0pt')};
catch
end
p.WhiteSpace = 'preserve';
% Use paragraph format to create vertical space.
try
    p.Format.SpaceAfter = sprintf('%gpt', nPts);
catch
    % Older DOM versions may not expose Format; fallback to an empty text run.
    append(p, mlreportgen.dom.Text(' '));
end
end

% end