package alekseybykov.portfolio.apache.poiooxml.consts;

public enum Consts {

	TAB_NAME("TAB_NAME");

	Consts(final String name) {
		this.name = name;
	}

	private final String name;

	@Override
	public String toString() {
		return name;
	}
}
