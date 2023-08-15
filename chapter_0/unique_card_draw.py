import random
suits = ["Clubs", "Diamonds", "Hearts", "Spades"]
faces = ["Jack", "Queen", "King", "Ace"]
numbered = [2, 3, 4, 5, 6, 7, 8, 9, 10]

deck = set()
for suit in suits:
    for card in faces + numbered:
        deck.add(f"{card} of {suit}")

def draw():
    card = random.choice(list(deck))
    deck.remove(card)
    return card

while (len(deck) > 0):
    if '2 of Diamonds' not in deck:
        print("2 of Diamonds is missing!")
    else:
        print("2 of Diamonds is still in the deck.")
    print(draw())

