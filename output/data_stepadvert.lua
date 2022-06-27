local KT = {id = 1,cost = 2,items = 3,step = 4}
local data = { 
 	[1] = {1,{{2,900}},{},5},
	[2] = {2,{{2,1900}},{{20001,1}},5},
	[3] = {3,{{2,2900}},{{20001,1},{20002,1}},5},
	[4] = {4,{{2,3900}},{{20001,1},{20002,1},{20003,1}},5}
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