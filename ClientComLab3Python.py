# -*- coding: UTF-8 -*-
import win32com.client

def main():
    try:
        obj_fn = win32com.client.Dispatch("Lb34MyFn.1")
        print(f"Объект {str(obj_fn)} получен.\n")
    except Exception as e:
        print(f"Ошибка при получении объекта: {e}")
        exit()

    num1 = 1
    num2 = 2
    num3 = 3

    result_fun141: float = obj_fn.Fun141(num1, num2)
    print(f"Результат вызова метода Fun141: {result_fun141}")

    result_fun142: int = obj_fn.Fun142(num1, num2, num3)
    print(f"Результат вызова метода Fun142: {result_fun142}")

    num4 = 1.5
    result_fun143: float = obj_fn.Fun143(num4)
    print(f"Результат вызова метода Fun143: {result_fun143}")

    return 0



if __name__ == "__main__":
    main()