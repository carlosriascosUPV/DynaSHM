function channels = extractRawChannels(data)

fields = fieldnames(data);
channels = struct('name',{},'signal',{},'time',{});

for i = 1:numel(fields)
    name = fields{i};

    if endsWith(name,'_TIME') || contains(name,'_AVG') || ...
       contains(name,'_MAX') || contains(name,'_MIN') || ...
       contains(name,'_RMS')
        continue
    end

    timeField = [name '_TIME'];
    if isfield(data,timeField)
        ch.name   = name;
        ch.signal = data.(name);
        ch.time   = data.(timeField);
        channels(end+1) = ch; %#ok<AGROW>
    end
end

if isempty(channels)
    error('No valid channels found.');
end
end
