package alekseybykov.portfolio.apache.poiooxml.model;

public class BookBuilder implements Builder {

	private Book book = new Book();

	@Override
	public Builder setIsbn(String isbn) {
		book.isbn = isbn;
		return this;
	}

	@Override
	public Builder setTitle(String title) {
		book.title = title;
		return this;
	}

	@Override
	public Builder setAuthor(String author) {
		book.author = author;
		return this;
	}

	@Override
	public Builder setPublisher(String publisher) {
		book.publisher = publisher;
		return this;
	}

	@Override
	public Builder setPrice(float price) {
		book.price = price;
		return this;
	}

	@Override
	public Book build() {
		return book;
	}
}
