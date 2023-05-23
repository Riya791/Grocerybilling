"""This is the model containing classes"""
from pydantic import BaseModel


class Item(BaseModel):
    id: int
    name: str
    quantity: int
    cost: int


class Email(BaseModel):
    rec_email: str
    subject: str
    body: str
