﻿dir 'C:\Users\Public\Documents\CEMS access\Richard' | gci | rename-item -newname { $_.name -replace 'Bert','Richard' }