package alekseybykov.portfolio.apache.poi.officeopen.model;

import java.util.Objects;

/**
 * @author Aleksey Bykov
 * @since 28.05.2020
 */
public class Book {
	private String isbn;
	private String title;
	private String author;
	private String publisher;
	private int price;

	public String getIsbn() {
		return isbn;
	}

	public void setIsbn(String isbn) {
		this.isbn = isbn;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public String getPublisher() {
		return publisher;
	}

	public void setPublisher(String publisher) {
		this.publisher = publisher;
	}

	public int getPrice() {
		return price;
	}

	public void setPrice(int price) {
		this.price = price;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		Book book = (Book) o;
		return price == book.price &&
		       isbn.equals(book.isbn) &&
		       title.equals(book.title) &&
		       Objects.equals(author, book.author) &&
		       Objects.equals(publisher, book.publisher);
	}

	@Override
	public int hashCode() {
		return Objects.hash(isbn, title, author, publisher, price);
	}
}
