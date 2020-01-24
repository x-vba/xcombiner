"use strict";


/**
 * This function takes an array of VBA source code strings and combines them
 * all together, taking care to remove excess Option statements when combining
 * them. This assumes that the Option statements used are the same in all of 
 * the source code modules, as differing options will very likely result in the
 * code no longer working.
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {array} vbaSourceCodeArray is an array of all the VBA source codes from the modules as strings
 * @param {string} [combinedSourceCodeModuleName="combinedModule"] is an optional paramter to set the
 * is module name of the combined module, and if omitted will be named combinedModule 
 * @returns {string} the VBA source code from all modules combined
 */
function vbaCombiner(vbaSourceCodeArray, combinedSourceCodeModuleName="combinedModule") {
	
	// Combining all the source code strings into a single source code string
	vbaSourceCodeArray = vbaSourceCodeArray.join("\n");
	vbaSourceCodeArray = vbaSourceCodeArray.split("\n");
	
	// Removing all Module Attribute names found at the top of the code modules
	vbaSourceCodeArray = vbaSourceCodeArray.filter(codeLine => {
		if (codeLine.trim().startsWith(`Attribute VB_Name = "`)) {
			return false
		}
		
		return true
	});
	
	// Creating an Array of all Option statements in the source code
	let allOptionStatements = [];
	vbaSourceCodeArray.forEach(codeLine => {
		if (codeLine.trim().toLowerCase().startsWith("option ")) {
			allOptionStatements.push(codeLine.trim());
		}
	});
	
	// Removing all Option statements from the combined source code
	vbaSourceCodeArray = vbaSourceCodeArray.filter(codeLine => {
		if (allOptionStatements.includes(codeLine.trim())) {
			return false
		}
		
		return true
	});
	
	// Removing duplicate options
	allOptionStatements = Array.from(new Set(allOptionStatements));
	
	// Adding the new Code Module Name at the top of the combined source code
	allOptionStatements.unshift(`Attribute VB_Name = "${combinedSourceCodeModuleName}"`);
	
	// Adding all Option statements and the new Code Module name to the top
	// of the source code
	vbaSourceCodeArray = allOptionStatements.concat(vbaSourceCodeArray);
	
	return vbaSourceCodeArray.join("\n")
}


module.exports = {
	vbaCombiner: vbaCombiner,	
};
