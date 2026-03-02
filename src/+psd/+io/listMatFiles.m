function mats = listMatFiles(inputPath)

mats = {};

if isfolder(inputPath)
    d = dir(fullfile(inputPath,'*.mat'));
    mats = cell(numel(d),1);
    for i = 1:numel(d)
        mats{i} = fullfile(d(i).folder, d(i).name);
    end
    return
end

if isfile(inputPath) && endsWith(lower(inputPath),'.mat')
    mats = {inputPath};
    return
end

error('Invalid input path.');
end
