function [tStart, tEnd] = timeCropGUI(t, y)

f = figure('Name','PSD Time Crop','Position',[300 300 980 420]);

ax = axes(f,'Position',[0.07 0.34 0.90 0.62]);
hPlot = plot(ax, t, y,'k'); grid on;
title(ax,'Reference Channel','FontName','Times New Roman','FontSize',14);

s1 = uicontrol(f,'Style','slider','Units','normalized', ...
    'Position',[0.10 0.20 0.35 0.05],'Min',t(1),'Max',t(end),'Value',t(1));
s2 = uicontrol(f,'Style','slider','Units','normalized', ...
    'Position',[0.55 0.20 0.35 0.05],'Min',t(1),'Max',t(end),'Value',t(end));

uicontrol(f,'Style','pushbutton','String','OK','Units','normalized', ...
    'Position',[0.45 0.06 0.10 0.09],'FontWeight','bold', ...
    'Callback','uiresume(gcbf)');

addlistener(s1,'Value','PostSet',@updatePlot);
addlistener(s2,'Value','PostSet',@updatePlot);

uiwait(f);

tStart = min(s1.Value, s2.Value);
tEnd   = max(s1.Value, s2.Value);

if isvalid(f); close(f); end

function updatePlot(~,~)
    t1 = min(s1.Value, s2.Value);
    t2 = max(s1.Value, s2.Value);
    idx = t >= t1 & t <= t2;
    set(hPlot,'XData',t(idx),'YData',y(idx));
    drawnow;
end
end
