using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using static CardObj;

// This script handles the creation of Deck objects.
public class DeckObj : MonoBehaviour
{
    public class Deck
    {
        // Each Deck will be a HashTable of Cards.
        public Hashtable Cards = new Hashtable();
        // Each Deck will have a specific size.
        public int size;
        // 4 suits in the deck
        public int numSuits = 4;
        // 13 cards per suit
        public int cardsPerSuit = 13;

        // Constructor with specific size.
        public Deck (int size)
        {
            this.size = size;
        }

        // Constructor with no initial size.
        public Deck(){}

        // Get and Set functions for the size.
        public int getSize()
        {
            return size;
        }
        public void setSize(int newSize)
        {
            size = newSize;
        }
        // Get and Set functions for the Cards Hashtable.
        public Hashtable getCards()
        {
            return Cards;
        }
        public void setCards(Hashtable newCards)
        {
            Cards = newCards;
        }

        // This function fills out the Hashtable of cards.
        public void initializeDeck()
        {
            // Iterate through each suit
            // int i denotes Suit
            for(int i = 1; i <= numSuits; i++)
            {
                // Iterate through each card per suit
                // int j denotes the value of the specific card
                for(int j = 1; j <= cardsPerSuit; j++)
                {
                    // Create a new card of suit i, and number j
                    // Using the Card(int, int) Constructor.
                    // The suit is converted to String within the Constructor.
                    Card newCard = new Card(i, j);
                    // Set the Key value for the Hashtable.
                    int key = j + (cardsPerSuit * (i - 1));
                    // Assign the newCard to the Deck using its Hashtable.
                    Cards.Add(key, newCard);
                }
            }

        }
    }
}
