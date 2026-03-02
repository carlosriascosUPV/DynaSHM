function channelsCut = cutAllChannels(channels, tStart, tEnd)

channelsCut = channels;

for i = 1:numel(channels)
    t = channels(i).time;
    idx = t >= tStart & t <= tEnd;
    channelsCut(i).time   = t(idx);
    channelsCut(i).signal = channels(i).signal(idx);
end
end
