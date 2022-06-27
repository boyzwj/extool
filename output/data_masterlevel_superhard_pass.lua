local KT = {id = 1,gold = 2,diamond = 3,editorId = 4,type = 5,background = 6,saveMinTimes = 7,saveSuccess = 8,saveFail = 9,saveMax = 10,oneStar = 11,twoStar = 12,threeStar = 13,maxPower = 14,reward = 15}
local data = { 
 	[1] = {1,50,25,305,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[2] = {2,50,25,160,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[3] = {3,50,25,170,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[4] = {4,50,25,200,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[5] = {5,50,25,220,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[6] = {6,50,25,230,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[7] = {7,50,25,290,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[8] = {8,50,25,310,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[9] = {9,50,25,312,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}},
	[10] = {10,50,25,317,3,"puzzle_bg_05",0,60,10,100,nil,nil,nil,100,{}}
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