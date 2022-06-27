local KT = {channel_id = 1,server_id = 2}
local data = { 
 	["96_3"] = {"96_3",4},
	["101_5"] = {"101_5",4},
	["96_5"] = {"96_5",18},
	["96_6"] = {"96_6",18},
	["96_7"] = {"96_7",18}
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