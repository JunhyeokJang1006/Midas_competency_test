str = '(1) test';
tokens = regexp(str, '\\((\\d+)\\)');
disp(tokens);
str2 = '(1) test';
tokens2 = regexp(str2, '\((\d+)\)');
disp(tokens2);
