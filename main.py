from sys import platform

if platform == 'darwin':
	import Excel_Splitting_V3_Mac
	Excel_Splitting_V3_Mac.main()
else:
	import Excel_Splitting_V3_Win
	Excel_Splitting_V3_Win.main()