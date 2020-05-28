package alekseybykov.portfolio.apache.poiooxml.model;

import java.util.Objects;

public class Book {
	protected String isbn;
	protected String title;
	protected String author;
	protected String publisher;
	protected float price;

	public String getIsbn() {
		return isbn;
	}

	public String getTitle() {
		return title;
	}

	public String getAuthor() {
		return author;
	}

	public String getPublisher() {
		return publisher;
	}

	public float getPrice() {
		return price;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Book book = (Book) o;
		return Float.compare(book.price, price) == 0 &&
		       Objects.equals(isbn, book.isbn) &&
		       Objects.equals(title, book.title) &&
		       Objects.equals(author, book.author) &&
		       Objects.equals(publisher, book.publisher);
	}

	@Override
	public int hashCode() {
		return Objects.hash(isbn, title, author, publisher, price);
	}
}
