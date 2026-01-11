import pytest
from retail.inventory import Inventory


class TestInventory:
    def test_add_item(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)

        item = inv.get_item("apple")
        assert item["quantity"] == 10
        assert item["price"] == 1.50

    def test_add_item_increases_quantity(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)
        inv.add_item("apple", 5, 1.50)

        item = inv.get_item("apple")
        assert item["quantity"] == 15

    def test_add_item_negative_quantity_raises(self):
        inv = Inventory()
        with pytest.raises(ValueError, match="Quantity cannot be negative"):
            inv.add_item("apple", -5, 1.50)

    def test_add_item_negative_price_raises(self):
        inv = Inventory()
        with pytest.raises(ValueError, match="Price cannot be negative"):
            inv.add_item("apple", 10, -1.50)

    def test_remove_item(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)
        inv.remove_item("apple", 3)

        item = inv.get_item("apple")
        assert item["quantity"] == 7

    def test_remove_item_deletes_when_zero(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)
        inv.remove_item("apple", 10)

        assert "apple" not in inv.list_items()

    def test_remove_item_not_found_raises(self):
        inv = Inventory()
        with pytest.raises(KeyError, match="not found"):
            inv.remove_item("apple", 5)

    def test_remove_item_too_many_raises(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)
        with pytest.raises(ValueError, match="Cannot remove more"):
            inv.remove_item("apple", 15)

    def test_get_item_not_found_raises(self):
        inv = Inventory()
        with pytest.raises(KeyError, match="not found"):
            inv.get_item("apple")

    def test_get_total_value(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)  # 15.00
        inv.add_item("banana", 5, 0.75)  # 3.75

        assert inv.get_total_value() == 18.75

    def test_get_total_value_empty(self):
        inv = Inventory()
        assert inv.get_total_value() == 0.0

    def test_list_items(self):
        inv = Inventory()
        inv.add_item("apple", 10, 1.50)
        inv.add_item("banana", 5, 0.75)

        items = inv.list_items()
        assert "apple" in items
        assert "banana" in items
        assert len(items) == 2
