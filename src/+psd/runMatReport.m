function out = runMatReport(inputPath, opts)

if nargin < 2 || isempty(opts)
    opts = struct();
end
opts = psd.local.applyDefaults(opts);

mats = psd.io.listMatFiles(inputPath);
if isempty(mats)
    error('No MAT files found.');
end

out = struct('matPath',{},'baseFolder',{},'reportPath',{});

for i = 1:numel(mats)

    matPath = mats{i};
    [matFolder, matName] = fileparts(matPath);

    baseFolder = fullfile(matFolder, matName);
    if ~exist(baseFolder,'dir')
        mkdir(baseFolder);
    end

    data = psd.io.loadMat(matPath);
    channels = psd.channels.extractRawChannels(data);

    % Un solo cuadro de diálogo: canal de referencia + opciones de reporte
    [refName, opts] = psd.ui.reportOptionsGUI({channels.name}, opts);
    idxRef = strcmp({channels.name}, refName);

    [tStart, tEnd] = psd.ui.timeCropGUI( ...
        channels(idxRef).time, ...
        channels(idxRef).signal);

    channelsCut = psd.channels.cutAllChannels(channels, tStart, tEnd);

    % ==========================================================
    % 1) Generar imágenes (PNG) primero, bien bonitas
    % ==========================================================
    imgPaths = psd.plot.plotAllChannels(channelsCut, baseFolder, opts);

    % ==========================================================
    % 2) Armar Word a partir de esas imágenes
    % ==========================================================
    reportPath = '';
    if opts.ExportWord
        reportPath = psd.report.exportWordFromFigures(imgPaths, baseFolder, opts);
    end

    out(end+1) = struct( ...
        'matPath', matPath, ...
        'baseFolder', baseFolder, ...
        'reportPath', reportPath); %#ok<AGROW>
end
end