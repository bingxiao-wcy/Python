ten_things = "Apples Oranges Crows Telephone Light Sugar"

print("Wait there's not 10 things in that list, let's fix that")

stuff = ten_things.split(' ')
more_stuff = ["Day","Night","Song","Frisbee","Corn","Banana","Girl","Boy"]

while len(stuff) != 10:
    next_one = more_stuff.pop()
    stuff.append(next_one)
    print("There's %d items now." % len(stuff))

print("There we go: ",stuff)

print("Let's do some things with stuff.")

print(stuff[1])# get first element
print(stuff[-1])# get last element
print(stuff.pop())# get last element
print(' '.join(stuff))
print('#'.join(stuff[3:5]))
