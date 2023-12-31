# %%
suits = ["Clubs", "Diamonds", "Hearts", "Spades"]
faces = ["Jack", "Queen", "King", "Ace"]
numbered = [2, 3, 4, 5, 6, 7, 8, 9, 10]


# %%
import random

# %%
def draw():
    which_suit = random.choice(suits)
    which_type = random.choice([faces, numbered])
    which_card = random.choice(which_type)
    return which_card, "of", which_suit


# %%
draw()

# %%
draw()

# %%
for n in range(5):
    card = draw()
    print(f"{n} - {card}")
    

# %%
deck = set()
for suit in suits:
    for card in faces + numbered:
        deck.add(f"{card} of {suit}")

# %%
len(deck)

# %%
print(deck)

# %%
print(sorted(deck))

# %%
print(dir(deck))

# %%
help(deck.remove)

# %%
card = random.choice(list(deck))
print(card)

# %%
deck.remove(card)
len(deck)

# %%
type(card)

# %%



