cities = {'CA':'San Francisco','MI':'Detroit','FL':'Jacksonville'}

cities['NY'] = 'New York'
cities['OR'] = 'Portland'

def find_city(themap,state):
    if state in themap:
        return themap[state]
    else:
        return "Not found"

#ok, pay attention
cities['_find'] = find_city# give a function to a dict variable 

while True:
    print("State?(ENTER to quit)"),
    state = input(">")
    if not state: break

    #this line is the most important ever! study!
    city_found = cities['_find'](cities,state)# through using find_city() function can get same result 
    print(city_found)
