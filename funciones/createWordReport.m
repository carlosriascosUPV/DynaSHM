function reportPath = createWordReport( ...
    inputFolder, exportFolder, packageRows, packageCols, ...
    startFigure, startPage, templateDocx, dpi, marginMm, factor)

% createWordReport (SVG + PNG reference crop)
% - Prioridad:
%     1) EMF (vector real en Word)
%     2) SVG recortado usando PNG como referencia (mantiene vector)
%     3) PNG (crop 1mm)
%     4) JPG/TIF/BMP (crop 1mm -> PNG temporal)
%
% - Paquetes 1x1,1x2,2x1,2x2 con letras a) b) c) d)
% - Caption: Figure N. a) Nombre1 + b) Nombre2 ...
% - Espacios:
%     * 6pt antes y después de cada paquete
%     * caption con 6pt antes y 6pt después
%     * dentro de tabla 0pt
% - Saltos:
%     * 1x2 o 2x1 -> 2 paquetes por página
%     * 2x2 o 1x1 -> 1 paquete por página

%% ==========================
% Defaults por nargin
% ==========================
if nargin < 1 || isempty(inputFolder),  inputFolder  = pwd; end
if nargin < 2 || isempty(exportFolder), exportFolder = fullfile(inputFolder,'REPORT_OUT'); end
if nargin < 3 || isempty(packageRows),  packageRows  = 1; end
if nargin < 4 || isempty(packageCols),  packageCols  = 2; end
if nargin < 5 || isempty(startFigure),  startFigure  = 1; end
if nargin < 6 || isempty(startPage),    startPage    = 1; end %#ok<NASGU>
if nargin < 7 || isempty(templateDocx), templateDocx = ""; end
if nargin < 8 || isempty(dpi),          dpi          = 600; end
if nargin < 9 || isempty(marginMm),     marginMm     = 1; end
if nargin < 10 || isempty(factor),      factor       = 0.45; end

packageRows = max(1, round(packageRows));
packageCols = max(1, round(packageCols));
if packageRows > 2, packageRows = 2; end
if packageCols > 2, packageCols = 2; end

imgsPerPackage = packageRows * packageCols;

% Paquetes por página según tu regla
if (packageRows==1 && packageCols==2) || (packageRows==2 && packageCols==1)
    packagesPerPage = 2;
else
    packagesPerPage = 1;
end

if ~exist(exportFolder,'dir')
    mkdir(exportFolder);
end

%% ==========================
% CONFIGURACIÓN “indispensable” (tu script original)
% ==========================
anchoPagina_cm = 21;
margen_cm      = 2.54;
anchoUtil_cm   = anchoPagina_cm - 2*margen_cm;

baseFactor = 0.45;
scale = factor / baseFactor;

if packageCols == 2
    imgWidth_cm = (anchoUtil_cm/2) * scale;
else
    imgWidth_cm = anchoUtil_cm * min(scale,1.0);
end
imgWidth_cm = min(imgWidth_cm, anchoUtil_cm);

%% ==========================
% LEER ARCHIVOS (incluye SVG y PNG)
% - luego nos quedamos con 1 por basename con prioridad:
%     EMF > SVG > PNG > JPG/JPEG > TIF/TIFF > BMP
% ==========================================================
exts = {'*.svg','*.png','*.jpg','*.jpeg','*.tif','*.tiff','*.bmp'};
files = [];
for e = 1:numel(exts)
    files = [files; dir(fullfile(inputFolder, exts{e}))]; %#ok<AGROW>
end
if isempty(files)
    error(['No se encontraron imágenes en: ', char(inputFolder)]);
end

[~,ix] = sort({files.name});
files = files(ix);

files = preferBestFormat_withSVG(files);
Nfigs = numel(files);
if Nfigs == 0
    error('No quedaron archivos después de filtrar por formato preferido.');
end

%% ==========================
% Word
% ==========================
import mlreportgen.dom.*

fname = fullfile(exportFolder, ['Report (', datestr(now,'dd-mm-yyyy HH-MM'), ')']);

if strlength(string(templateDocx)) > 0 && exist(templateDocx,'file') == 2
    doc = Document(fname, 'docx', templateDocx);
else
    doc = Document(fname, 'docx');
end

letters = 'abcdefghijklmnopqrstuvwxyz';

%% ==========================
% Helpers de estilo
% ==========================
    function p = spacer(npt)
        p = Paragraph();
        p.WhiteSpace = 'preserve';
        p.Style = {OuterMargin('0pt','0pt',sprintf('%dpt',npt),'0pt')};
        append(p, Text(' '));
    end

    function p = labelABCD(idxLetter)
        p = Paragraph();
        p.Style = {HAlign('center'), OuterMargin('0pt','0pt','0pt','0pt'), ...
                   FontFamily('Times New Roman'), FontSize('13pt')};
        t1 = Text(letters(idxLetter)); t1.Italic = true;
        t2 = Text(')');
        append(p, t1);
        append(p, t2);
    end

    function p = captionPackage(~, names, useLetters)

    letters = 'abcdefghijklmnopqrstuvwxyz';

    p = Paragraph();
    p.Style = { ...
        FontFamily('Times New Roman'), ...
        FontSize('13pt'), ...
        HAlign('center'), ...
        OuterMargin('6pt','0pt','6pt','0pt') ...
        };

    nbsp = char(160); % espacio no separable (Word lo respeta siempre)

    % "Figure . "
    append(p, Text(['Figure', nbsp, '.', nbsp]));

    for ii = 1:numel(names)

        if useLetters
            tLetter = Text(letters(ii));
            tLetter.Italic = true;
            append(p, tLetter);

            append(p, Text([')', nbsp]));
        end

        append(p, Text(extractAfter(names{ii},'. ')));

        if ii < numel(names)
            append(p, Text([';', nbsp]));
        end

    end

end

%% ==========================
% LOOP PRINCIPAL (por paquetes)
% ==========================
k = 1;
pkgIdx = 0;

while k <= Nfigs
    pkgIdx = pkgIdx + 1;
    figNum = startFigure + (pkgIdx - 1);

    kEnd = min(Nfigs, k + imgsPerPackage - 1);
    pkgFiles = files(k:kEnd);
    nThis = numel(pkgFiles);

    useLetters = (nThis > 1);

    append(doc, spacer(6)); % 6pt antes del paquete

    t = Table();
    t.Style = {Width('100%')};
    t.Border = 'none';

    imgIdx = 1;
    names = cell(1,nThis);
    usedCount = 0;

    for r = 1:packageRows
        rowImg = TableRow();
        rowLbl = TableRow();

        for c = 1:packageCols
            cellImg = TableEntry();
            cellLbl = TableEntry();

            cellImg.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
            cellLbl.Style = {OuterMargin('0pt','0pt','0pt','0pt')};

            if imgIdx <= nThis
                imgPath = fullfile(inputFolder, pkgFiles(imgIdx).name);

                % ✅ nuevo resolver: si hay SVG y PNG hermano, recorta SVG usando PNG
                [insertPath, baseNameNoExt, insertAsVector, refPngForAR] = resolveForWord_EMF_or_CroppedSVG_orPNG(imgPath, dpi, marginMm);

                if ~isempty(insertPath)
                    usedCount = usedCount + 1;
                    names{usedCount} = baseNameNoExt;

                    img = Image(insertPath);

                    if insertAsVector
                        % EMF o SVG: solo width
                        img.Width = sprintf('%.2fcm', imgWidth_cm);
                    else
                        % raster: ancho + alto por aspect ratio
                        info = imfinfo(insertPath);
                        imgHeight_cm = (info.Height / info.Width) * imgWidth_cm;
                        img.Width  = sprintf('%.2fcm', imgWidth_cm);
                        img.Height = sprintf('%.2fcm', imgHeight_cm);
                    end

                    % Si insertamos SVG (vector) y tenemos PNG referencia,
                    % usamos el PNG para fijar el alto sin deformar (opcional robusto)
                    % pero OJO: Word suele respetar width-only mejor para vector
                    %#ok<NASGU>

                    pImg = Paragraph();
                    pImg.Style = {HAlign('center'), OuterMargin('0pt','0pt','0pt','0pt')};
                    append(pImg, img);
                    append(cellImg, pImg);

                    if useLetters
                        append(cellLbl, labelABCD(usedCount));
                    else
                        pEmpty = Paragraph(); pEmpty.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
                        append(pEmpty, Text(' '));
                        append(cellLbl, pEmpty);
                    end
                else
                    pEmpty = Paragraph(); pEmpty.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
                    append(pEmpty, Text(' '));
                    append(cellImg, pEmpty);

                    pEmpty2 = Paragraph(); pEmpty2.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
                    append(pEmpty2, Text(' '));
                    append(cellLbl, pEmpty2);
                end
            else
                pEmpty = Paragraph(); pEmpty.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
                append(pEmpty, Text(' '));
                append(cellImg, pEmpty);

                pEmpty2 = Paragraph(); pEmpty2.Style = {OuterMargin('0pt','0pt','0pt','0pt')};
                append(pEmpty2, Text(' '));
                append(cellLbl, pEmpty2);
            end

            append(rowImg, cellImg);
            append(rowLbl, cellLbl);

            imgIdx = imgIdx + 1;
        end

        append(t, rowImg);
        append(t, rowLbl);
    end

    append(doc, t);

    if usedCount > 0
        cap = captionPackage(figNum, names(1:usedCount), useLetters);
        append(doc, cap);
    else
        warning('Paquete %d quedó vacío (no se insertó ninguna imagen).', pkgIdx);
    end

    append(doc, spacer(6)); % 6pt después del paquete

    if mod(pkgIdx, packagesPerPage) == 0
        append(doc, PageBreak());
    end

    k = kEnd + 1;
end

close(doc);

reportPath = char(fname) + ".docx";
disp("Documento Word generado: " + reportPath);

save(fullfile(exportFolder,'name.mat'),'reportPath')

end


%% ========================================================================
% Preferir mejor formato por base name (CON SVG)
% Prioridad: emf > svg > png > jpg/jpeg > tif/tiff > bmp
% ========================================================================
function filesOut = preferBestFormat_withSVG(filesIn)

prio = containers.Map( ...
    {'.tit','.svg','.png','.jpg','.jpeg','.tif','.tiff','.bmp'}, ...
    [  1,    2,    3,    4,     4,     5,     5,     6 ]);

bestIdx = containers.Map();

for i = 1:numel(filesIn)
    [~,bn,ext] = fileparts(filesIn(i).name);
    ext = lower(ext);
    if ~isKey(prio, ext), continue; end

    if ~isKey(bestIdx, bn)
        bestIdx(bn) = i;
    else
        j = bestIdx(bn);
        [~,~,extJ] = fileparts(filesIn(j).name);
        extJ = lower(extJ);
        if prio(ext) < prio(extJ)
            bestIdx(bn) = i;
        end
    end
end

idx = cell2mat(values(bestIdx));
idx = sort(idx);
filesOut = filesIn(idx);

end


%% ========================================================================
% Resolver final:
% - si hay EMF hermano -> usarlo
% - si es SVG y hay PNG hermano -> crear SVG recortado (viewBox) usando PNG y usarlo
% - si hay PNG -> crop 1mm y usar PNG
% - si raster -> crop 1mm a PNG temporal
% ========================================================================
function [insertPath, baseNameNoExt, insertAsVector, refPng] = resolveForWord_EMF_or_CroppedSVG_orPNG(imgPath, dpi, marginMm)

[folder, baseNameNoExt, ext] = fileparts(imgPath);
ext = lower(ext);

emf = [];%fullfile(folder, [baseNameNoExt,'.emf']);
svg = fullfile(folder, [baseNameNoExt,'.svg']);
png = fullfile(folder, [baseNameNoExt,'.png']);

refPng = '';
insertAsVector = false;

% 1) EMF siempre manda
if exist(emf,'file') == 2
    insertPath = emf;
    insertAsVector = true;
    return
end

% 2) SVG con PNG referencia -> recortar SVG (vector) y usarlo
if exist(svg,'file') == 2 && exist(png,'file') == 2
    refPng = png;

    % Crear SVG recortado en temp (para no tocar el original)
    svgCrop = fullfile(tempdir, ['svgcrop_' baseNameNoExt '_' char(java.util.UUID.randomUUID) '.svg']);

    try
        cropSVG_usingPNG(svg, png, svgCrop, dpi, marginMm);
        insertPath = svgCrop;
        insertAsVector = true;
        return
    catch
        % si falla, caemos a PNG
    end
end

% 3) PNG (si existe) -> crop 1mm y usarlo
if exist(png,'file') == 2
    insertPath = cropRasterToTempPNG_1mm(png, dpi, marginMm, baseNameNoExt);
    insertAsVector = false;
    return
end

% 4) Raster genérico
if any(strcmp(ext,{'.jpg','.jpeg','.tif','.tiff','.bmp','.png'}))
    insertPath = cropRasterToTempPNG_1mm(imgPath, dpi, marginMm, baseNameNoExt);
    insertAsVector = false;
    return
end

% 5) SVG sin PNG -> no sabemos crop, insertamos SVG tal cual (opcional)
if strcmp(ext,'.svg') && exist(svg,'file')==2
    insertPath = svg;          % sin crop
    insertAsVector = true;
    return
end

insertPath = '';

end


%% ========================================================================
% Crop raster a PNG temporal con borde 1 mm exacto
% ========================================================================
function tmpName = cropRasterToTempPNG_1mm(rasterPath, dpi, marginMm, baseNameNoExt)

I = imread(rasterPath);

if size(I,3) == 3
    G = rgb2gray(I);
else
    G = I;
end

thr  = 250;
mask = (G < thr);
mask = bwareaopen(mask, 50);

if any(mask(:))
    [r,c] = find(mask);
    r1 = min(r); r2 = max(r);
    c1 = min(c); c2 = max(c);

    pad = round((marginMm/25.4) * dpi);

    r1 = max(1, r1 - pad);
    r2 = min(size(I,1), r2 + pad);
    c1 = max(1, c1 - pad);
    c2 = min(size(I,2), c2 + pad);

    Icrop = I(r1:r2, c1:c2, :);
else
    Icrop = I;
end

tmpName = fullfile(tempdir, ['crop_' baseNameNoExt '_' char(java.util.UUID.randomUUID) '.png']);
imwrite(Icrop, tmpName);

end


%% ========================================================================
% Recorta SVG (vector) ajustando viewBox usando bbox medido en PNG
% ========================================================================
function cropSVG_usingPNG(svgIn, pngRef, svgOut, dpi, marginMm)

if nargin < 4 || isempty(dpi), dpi = 600; end
if nargin < 5 || isempty(marginMm), marginMm = 1; end

I = imread(pngRef);

if size(I,3)==3, G = rgb2gray(I); else, G = I; end

thr = 250;
mask = (G < thr);
mask = bwareaopen(mask, 50);

if ~any(mask(:))
    error('PNG referencia sin contenido detectable. Ajusta umbral.');
end

[r,c] = find(mask);
top    = min(r); bottom = max(r);
left   = min(c); right  = max(c);

pad = round((marginMm/25.4)*dpi);

top    = max(1, top-pad);
bottom = min(size(I,1), bottom+pad);
left   = max(1, left-pad);
right  = min(size(I,2), right+pad);

pxW = size(I,2); pxH = size(I,1);
bboxPx = [left, top, right-left+1, bottom-top+1];

txt = fileread(svgIn);

m = regexp(txt, 'viewBox\s*=\s*"([^"]+)"', 'tokens', 'once');
if isempty(m)
    error('SVG sin viewBox: no puedo recortar de forma segura.');
end
vb = sscanf(m{1}, '%f %f %f %f');
vbX = vb(1); vbY = vb(2); vbW = vb(3); vbH = vb(4);

sx = vbW / pxW;
sy = vbH / pxH;

newVbX = vbX + (bboxPx(1)-1) * sx;
newVbY = vbY + (bboxPx(2)-1) * sy;
newVbW = bboxPx(3) * sx;
newVbH = bboxPx(4) * sy;

newViewBoxStr = sprintf('viewBox="%.12g %.12g %.12g %.12g"', newVbX, newVbY, newVbW, newVbH);
txt = regexprep(txt, 'viewBox\s*=\s*"[^"]+"', newViewBoxStr, 'once');

% Ajustar width/height si existen
[wVal,wUnit] = parseSizeToken(regexp(txt,'width\s*=\s*"([^"]+)"','tokens','once'));
[hVal,hUnit] = parseSizeToken(regexp(txt,'height\s*=\s*"([^"]+)"','tokens','once'));

if ~isnan(wVal) && ~isnan(hVal)
    newW = wVal * (newVbW / vbW);
    newH = hVal * (newVbH / vbH);
    txt = regexprep(txt,'width\s*=\s*"[^"]+"', sprintf('width="%.12g%s"', newW, wUnit), 'once');
    txt = regexprep(txt,'height\s*=\s*"[^"]+"', sprintf('height="%.12g%s"', newH, hUnit), 'once');
end

fid = fopen(svgOut,'w');
fwrite(fid, txt);
fclose(fid);

end

function [val, unit] = parseSizeToken(tok)
if isempty(tok)
    val = NaN; unit = '';
    return
end
s = tok{1};
m = regexp(s,'^\s*([0-9]*\.?[0-9]+)\s*([a-zA-Z%]*)\s*$','tokens','once');
if isempty(m)
    val = NaN; unit = '';
    return
end
val = str2double(m{1});
unit = m{2};
end