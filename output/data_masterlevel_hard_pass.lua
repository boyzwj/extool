local KT = {id = 1,gold = 2,diamond = 3,editorId = 4,type = 5,background = 6,saveMinTimes = 7,saveSuccess = 8,saveFail = 9,saveMax = 10,oneStar = 11,twoStar = 12,threeStar = 13,maxPower = 14,reward = 15}
local data = { 
 	[1] = {1,50,25,155,2,"puzzle_bg_05",0,60,10,100,3800,6700,8700,100,{}},
	[2] = {2,50,25,175,2,"puzzle_bg_05",0,60,10,100,9000,15700,20100,100,{}},
	[3] = {3,50,25,185,2,"puzzle_bg_05",0,60,10,100,5300,8000,9900,100,{}},
	[4] = {4,50,25,195,2,"puzzle_bg_05",0,60,10,100,4800,8200,10700,100,{}},
	[5] = {5,50,25,215,2,"puzzle_bg_05",0,60,10,100,5000,9200,12200,100,{}},
	[6] = {6,50,25,225,2,"puzzle_bg_05",0,60,10,100,11000,16200,20000,100,{}},
	[7] = {7,50,25,235,2,"puzzle_bg_05",0,60,10,100,4000,9000,12700,100,{}},
	[8] = {8,50,25,265,2,"puzzle_bg_05",0,60,10,100,2100,4700,6700,100,{}},
	[9] = {9,50,25,275,2,"puzzle_bg_05",0,60,10,100,12400,19000,23800,100,{}},
	[10] = {10,50,25,295,2,"puzzle_bg_05",0,60,10,100,6800,11900,15600,100,{}}
}
do
	local base = {
		__index = function(table,key)
			local ki = KT[key]
			if not ki then
				return nil
			end
			return table[ki]
    	end,
		__newindex = function()
			error([[Attempt to modify read-only table]])
		end
	}
	for k, v in pairs(data) do
		setmetatable(v, base)
	end
	base.__metatable = false
end
return data