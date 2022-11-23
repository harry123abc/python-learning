# -*- coding: utf-8 -*-
"""
Created on Sun Mar 10 16:22:11 2019

@author: OVO
"""

class Animal(object):
    def run(self):
        print('Animal is running...')
        
class Dog(Animal):
    def run(self):
        print('Dog is running...')
    def eat(self):
        print('Eating meat...')

class Cat(Animal):
    def run(Animal):
        print('Cat is running...')

cat = Cat()
cat.run()

dog = Dog()
dog.run()

def run_twice(animal):
    animal.run()
    animal.run()
    
class Tortoise(Animal):
    def run(self):
        print('Tortoise is running slowly...')
        
class Timer(object):
    def run(self):
        print('Start...')