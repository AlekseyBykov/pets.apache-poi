package alekseybykov.portfolio.apache.poiooxml.model;

public interface Builder {

	Builder setIsbn(String isbn);

	Builder setTitle(String title);

	Builder setAuthor(String author);

	Builder setPublisher(String publisher);

	Builder setPrice(float price);

	Book build();
}
