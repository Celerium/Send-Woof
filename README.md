# Send-Woof

The Send-Woof script sends a dog image & fact to a Teams channel using a Teams webhook connector URI.

- Various filters are in place to try and prevent any inappropriate images from being sent.

An image is randomly selected from a random dog subreddit. Dog facts are pulled from the dog-facts-api.herokuapp.com API

---

## Send-Woof

![Send-Woof](https://raw.githubusercontent.com/Celerium/Send-Woof/main/.github/Celerium-Send-Woof-Example001.png)

## Initial Setup & Running

1. Teams Channel > Connectors > Incoming Webhook
2. Give the Webhook a name & logo
    - Create the Webhook
4. Copy the URI
    - The URI is how you tell the script what teams channel to send posts to.

---

```posh
    .\Send-Woof.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....'
```

Using the defined webhooks connector URI a random dad joke is sent to the webhooks Teams channel.

No output is displayed to the console.
Using the -Verbose option will give you a basic display output


## Help :blue_book:

  - Help info and a list of parameters can be found by running `Get-Help .\Send-Woof.ps1`, such as:

```posh
Get-Help .\Send-Woof.ps1
Get-Help .\Send-Woof.ps1 -Full
```

---