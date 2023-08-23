from frenchdeck import FrenchDeck

deck = FrenchDeck()

#card_one = deck.take_random_card()
#card_two = deck.take_random_card()
card_one = deck._cards[12]
card_two = deck._cards[11]
print(f"Card one: {card_one}")
print(f"Card two: {card_two}")

deck.holdem_start_hand_strength(card_one,card_two)

