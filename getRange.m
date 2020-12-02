function range = getRange(startR,startC,endR,endC,sheet)
    range = strcat(getCellCoords(startR,startC),":");
    range = strcat(range,getCellCoords(endR,endC));
    if exist('sheet','var')
        range = strcat(sheetCoord(sheet),range);
    end
end