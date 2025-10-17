try
    load('competency_correlation_workspace_20250915.mat','allData','periods');
catch
    error('워크스페이스 파일을 찾을 수 없습니다. 최신 파일 경로를 확인하세요.');
end

periodFields = fieldnames(allData);
for i = 1:length(periodFields)
    field = periodFields{i};
    if ~isfield(allData.(field),'selfData')
        continue;
    end
    selfData = allData.(field).selfData;
    varNames = selfData.Properties.VariableNames;
    fprintf('\n=== %s ===\n', field);
    for v = 1:length(varNames)
        name = varNames{v};
        col = selfData{:,v};
        if ~isnumeric(col)
            continue;
        end
        if ~(startsWith(name,'Q') || startsWith(name,'q'))
            continue;
        end
        vals = unique(col(~isnan(col)));
        if isempty(vals)
            continue;
        end
        if numel(vals) <= 15
            fprintf('%s: min %.0f, max %.0f, unique %s\n', name, min(vals), max(vals), mat2str(vals')); 
        else
            fprintf('%s: min %.2f, max %.2f (unique>15)\n', name, min(vals), max(vals));
        end
    end
end
