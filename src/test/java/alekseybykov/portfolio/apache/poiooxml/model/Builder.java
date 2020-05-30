package alekseybykov.portfolio.apache.poiooxml.model;

/**
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public interface Builder {

	Builder setIsbn(String isbn);

	Builder setTitle(String title);

	Builder setAuthor(String author);

	Builder setPublisher(String publisher);

	Builder setPrice(int price);

	Book build();
}
