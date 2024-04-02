using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using static CardObj;
using static DeckObj;

public class GameRunner : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        // Create a Deck
        Deck newDeck = new Deck();
        // Initialize the Deck
        newDeck.initializeDeck();
        // Get a specific card from the Deck.
        Hashtable currentCards = newDeck.getCards();
        Debug.Log(currentCards.GetHash(1).getValue());
        
    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
