package alekseybykov.portfolio.apache.poiooxml.model;

public interface Builder {

	public Builder setIsbn(String isbn);

	public Builder setTitle(String title);

	public Builder setAuthor(String author);

	public Builder setPublisher(String publisher);

	public Builder setPrice(float price);

	public Book build();
}
