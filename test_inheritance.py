from enum import Enum
import json
from typing import List

from pydantic import BaseModel, Field, model_validator


class OBJECT_TYPE(str, Enum):
    A = "A"
    B = "B"


class BaseObject(BaseModel):
    type: OBJECT_TYPE = Field(...)


class ObjectA(BaseObject):
    msg: str = Field(...)

    @model_validator(mode='before')
    def set_type(cls, values):
        values["type"] = OBJECT_TYPE.A
        return values


class ObjectB(BaseObject):
    value: int = Field(...)

    @model_validator(mode='before')
    def set_type(cls, values):
        values["type"] = OBJECT_TYPE.A
        return values


if __name__ == "__main__":
    a = ObjectA(msg="Hello")
    b = ObjectB(value=42)
    list: List[BaseObject] = [a, b]
    for obj in list:
        print(json.dumps(obj.model_dump()))
