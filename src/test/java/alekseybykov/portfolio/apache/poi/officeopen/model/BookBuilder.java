package alekseybykov.portfolio.apache.poi.officeopen.model;

/**
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public class BookBuilder implements Builder {

	private Book book = new Book();

	@Override
	public Builder setIsbn(String isbn) {
		book.setIsbn(isbn);
		return this;
	}

	@Override
	public Builder setTitle(String title) {
		book.setTitle(title);
		return this;
	}

	@Override
	public Builder setAuthor(String author) {
		book.setAuthor(author);
		return this;
	}

	@Override
	public Builder setPublisher(String publisher) {
		book.setPublisher(publisher);
		return this;
	}

	@Override
	public Builder setPrice(int price) {
		book.setPrice(price);
		return this;
	}

	@Override
	public Book build() {
		return book;
	}
}
