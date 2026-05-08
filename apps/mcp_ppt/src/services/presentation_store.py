from src.models.presentation import PresentationModel

class PresentationStore:
    def __init__(self) -> None:
        self._items: dict[str, PresentationModel] = {}

    def add(self, presentation: PresentationModel) -> None:
        self._items[presentation.id] = presentation

    def get(self, prs_id: str) -> PresentationModel:
        if prs_id not in self._items:
            raise ValueError(f"Presentation '{prs_id}' not found")

        return self._items[prs_id]

    def exists(self, prs_id: str) -> bool:
        return prs_id in self._items

    def remove(self, prs_id: str) -> None:
        self._items.pop(prs_id, None)