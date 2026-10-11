import random

import pandas as pd


def unique_phones(count: int) -> list[str]:
    """Genera 'count' números de teléfono de 10 dígitos, todos únicos."""
    phones = set()
    while len(phones) < count:
        phones.add(f"{random.randint(1_000_000_000, 9_999_999_999)}")
    return list(phones)


def main() -> None:
    n = 100

    # Hoja "data": filas de prueba con un flag active
    data_df = pd.DataFrame(
        {
            "id": list(range(1, n + 1)),
            "phone": unique_phones(n),
            "token": [f"{random.randint(0, 9999):04}" for _ in range(n)],
            "active": [1] * n,
        }
    )

    # Hoja "users": usuario de ejemplo
    users_df = pd.DataFrame(
        {
            "id": [1],
            "users": ["admin"],
            "password": ["admin"],
            "type": [1],
        }
    )

    with pd.ExcelWriter("data_prueba.xlsx", engine="openpyxl") as writer:
        data_df.to_excel(writer, sheet_name="data", index=False)
        users_df.to_excel(writer, sheet_name="users", index=False)

    print("Tablas de prueba generadas y guardadas en 'data_prueba.xlsx'")


if __name__ == "__main__":
    main()
