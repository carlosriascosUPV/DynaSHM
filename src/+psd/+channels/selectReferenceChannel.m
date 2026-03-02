function refChannel = selectReferenceChannel(channelNames)

[idx, ok] = listdlg( ...
    'PromptString','Select reference channel:', ...
    'SelectionMode','single', ...
    'ListString',channelNames, ...
    'ListSize',[320 360]);

if ~ok
    error('No channel selected.');
end

refChannel = channelNames{idx};
end
