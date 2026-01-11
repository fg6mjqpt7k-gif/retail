class Inventory:
    def __init__(self):
        self._items = {}

    def add_item(self, name: str, quantity: int, price: float) -> None:
        if quantity < 0:
            raise ValueError("Quantity cannot be negative")
        if price < 0:
            raise ValueError("Price cannot be negative")

        if name in self._items:
            self._items[name]["quantity"] += quantity
        else:
            self._items[name] = {"quantity": quantity, "price": price}

    def remove_item(self, name: str, quantity: int) -> None:
        if name not in self._items:
            raise KeyError(f"Item '{name}' not found in inventory")
        if quantity > self._items[name]["quantity"]:
            raise ValueError("Cannot remove more items than available")

        self._items[name]["quantity"] -= quantity
        if self._items[name]["quantity"] == 0:
            del self._items[name]

    def get_item(self, name: str) -> dict:
        if name not in self._items:
            raise KeyError(f"Item '{name}' not found in inventory")
        return self._items[name].copy()

    def get_total_value(self) -> float:
        return sum(
            item["quantity"] * item["price"]
            for item in self._items.values()
        )

    def list_items(self) -> list[str]:
        return list(self._items.keys())
