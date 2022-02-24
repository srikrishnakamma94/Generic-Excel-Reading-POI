package in.ainvnt.excel.parsing.enums;

public enum ExcelSection {

	ORDERS("Order"), //
	PROFIT("Profit");

	final String typeValue;

	private ExcelSection(final String typeValue) {
		this.typeValue = typeValue;
	}

	public String getName() {
		return name();
	}

	public String getValue() {
		return typeValue;
	}

	@Override
	public String toString() {
		return name();
	}

}
