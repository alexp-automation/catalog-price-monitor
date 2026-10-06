from bs4 import BeautifulSoup

from monitor import compare, parse_listing

LISTING = """
<article class="product_pod">
  <h3><a href="../../../a-light_1/index.html" title="A Light in the Attic">A Light...</a></h3>
  <div class="product_price">
    <p class="price_color">£51.77</p>
    <p class="instock availability"> In stock </p>
  </div>
  <p class="star-rating Three"></p>
</article>
<ul class="pager"><li class="next"><a href="page-2.html">next</a></li></ul>
"""


def test_parse_listing_reads_fields_and_next_page():
    url = "https://books.toscrape.com/catalogue/category/books/poetry_23/index.html"
    items, next_url = parse_listing(BeautifulSoup(LISTING, "html.parser"), url, "Poetry")
    assert items == [{
        "title": "A Light in the Attic",
        "category": "Poetry",
        "price": 51.77,
        "rating": 3,
        "in_stock": True,
        "url": "https://books.toscrape.com/catalogue/a-light_1/index.html",
    }]
    assert next_url.endswith("poetry_23/page-2.html")


def test_compare_finds_price_changes_and_new_items():
    old = [{"url": "a", "price": 10.0}, {"url": "b", "price": 5.0}]
    new = [{"url": "a", "price": 12.5}, {"url": "b", "price": 5.0}, {"url": "c", "price": 1.0}]
    changes, new_items = compare(new, old)
    assert [(c["url"], c["diff"]) for c in changes] == [("a", 2.5)]
    assert [p["url"] for p in new_items] == ["c"]
