local KT = {channel_id = 1,server_id = 2}
local data = { 
 	["96_3"] = {"96_3",{1,2,3,4,5,18}},
	["101_5"] = {"101_5",{1,2,3,4,5}},
	["96_5"] = {"96_5",{1,2,3}},
	["96_6"] = {"96_6",{1,2}},
	["96_7"] = {"96_7",{1,3}}
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