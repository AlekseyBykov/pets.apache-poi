package alekseybykov.portfolio.apache.poiooxml.consts;

public enum Consts {

	TAB_NAME("TAB_NAME"),

	ISBN_TITLE("ISBN"),
	BOOK_TITLE("BOOK TITLE"),
	AUTHOR_TITLE("AUTHOR"),
	PUBLISHER_TITLE("PUBLISHER"),
	PRICE_TITLE("PRICE");

	Consts(final String name) {
		this.name = name;
	}

	private final String name;

	@Override
	public String toString() {
		return name;
	}
}
