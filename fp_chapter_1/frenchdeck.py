import collections
import random
import holdem_starting_hand_strength

card = collections.namedtuple('Card', ['rank', 'suit'])

class FrenchDeck:
    ranks = [str(n) for n in range(2, 10)] + list('TJQKA')
    suits = 'spades diamonds clubs hearts'.split()

    def __init__(self):
        self._cards = [card(rank, suit) for suit in self.suits
                                        for rank in self.ranks]

    def __len__(self):
        return len(self._cards)

    def __getitem__(self, position):
        print(f"Type: {self._cards[position]}")
        return self._cards[position]    

    # Get a random card from the deck but it is not removed
    def get_random_card(self):
        return random.choice(self._cards)
    
    # Take a random card out of the deck
    def take_random_card(self):
        card = random.choice(self._cards)
        self._cards.remove(card)
        return card
    
    # Given two cards, determine their value compared to the other 168 combos
    def holdem_start_hand_strength(self, card_1, card_2):
        # Hand is paired
        if card_1.rank == card_2.rank:
            print(f"Paired")
            return holdem_starting_hand_strength.holdem_player_hand(
                rank=50, paired=True, suited=False, 
                combos=6, combo_percent=.45, win_percent=22)
        # Hand is not paired
        else:
            #Hand is suited
            if card_1.suit == card_2.suit:
                print(f"Suited: {self.highest_ranks(card_1, card_2)}")
                print(f"Value: {holdem_starting_hand_strength.suitedvalues(self.highest_ranks(card_1, card_2))}")
                return holdem_starting_hand_strength.holdem_player_hand(
                    rank=50, paired=False, suited=True, 
                    combos=4, combo_percent=.30, win_percent=22)
            #Hand is not suite
            else:
                print(f"Unsuited: {self.highest_ranks(card_1, card_2)}")
                print(f"Value: {holdem_starting_hand_strength.unsuitedvalues(self.highest_ranks(card_1, card_2))}")
                return holdem_starting_hand_strength.holdem_player_hand(
                    rank=50, paired=False, suited=False, 
                    combos=12, combo_percent=.90, win_percent=22)
             
    # Given two cards, we always want the highest card value first
    def highest_ranks(self, card_1, card_2):
        ranks = {'2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9, 'T': 10, 'J': 11, 'Q': 12, 'K': 13, 'A': 14}
        rank_1 = ranks[card_1.rank]
        rank_2 = ranks[card_2.rank]
        if rank_1 > rank_2:
            return(f"{card_1.rank}{card_2.rank}")
        else:
            return(f"{card_2.rank}{card_1.rank}")
    