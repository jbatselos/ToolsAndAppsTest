module.exports = {
	transform: {
		"^.+\\.tsx?$": "ts-jest",
	},
	testRegex: "(/__tests__/.*|(\\.|/)(test|spec))\\.(jsx?|ts)$",
	moduleFileExtensions: ["ts", "tsx", "js", "jsx", "json", "node"],
};
// testRegex: ".*\\src\\.*\.test\.(jsx?|tsx?|ts?|js)$",
  // testRegex: "(/src/.*|(\\.|/)(test|spec))\\.(jsx?|tsx?)$",
  // testPathIgnorePatterns: ["<rootDir>/src/Configurer", "<rootDir>/src/Enums"],
