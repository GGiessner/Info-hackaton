Le projet consiste en la rédaction d'une proposition commerciale pour un projet de conseil en stratégie.

Il s'agit de réaliser de manière complètement automatisée un contrat, en utilisant des notes prises pendant l'entretien, ainsi qu'un devis fait afin de répondre à la problématique du client. Nous avons également à notre disposition un Excel qui contient pleins de références d'anciennes missions, afin de construire un liste de références de missions réalisées similaires à la mission actuelle.
L'entreprise souhaite que nous utilisions des LLM afin de pousser aussi loin que possible l'automatisation de ce projet.


Pour effectuer cette mission proposée par l'entreprise GJOA, nous avons décidé d'écouter leurs conseils et d'utiliser des LLM.
Dans un premier temps, nous venons chercher les informations utiles au projet et diirectement utilisables telles qu'elles sont (le nom de la société, la date...).
Ensuite, nous utilisons les LLM pour traiter les notes prises pendant l'entretien. Cela nous permet de générer la contextualisation de la mission dans le contrat. Cela va également nous permettre de trouver des références de missions proches de ce que nous pouvons trouver dans le Excel des références similaires au projet en cours. Finalement, à partir de requêtes dans le devis, nous pouvons extraire les différents axes de travail du contrat, et nous pouvons donc, à partir d'une succinte description, générer un petit texte pour décrire les différentes sous-parties du programme.

Afin d'effectuer le contrat dans la plus grande confidentialité, nous faisons tourner le programme sur un programme de LLM en local. Cela évite toute potentielle fite d'information, mais rend le traitement plus long et potentiellement coûteux en termes de Hardware.

Pour ce projet, nous utilisons une librairie un peu exotique: Ollama






Première utilisation:


Réaliser l'installation de ollama via le site https://ollama.com/download (cliquer sur le .exe et effectuer toutes les procédures demandées pour l'installation)

Pour vérifier le bon fonctionnement de l'installation, mettre la commande "ollama run mistral" dans un terminal. Le terminal devrait indiquer un téléchargement, qui correspond au téléchargement du modèle d'IA de mistral



Utiliser le programme pour générer un contrat:


Aller dans un terminal, pour se placer dans le bon dossier avant d'ouvrir VS Code. La commande est "cd {chemin d'accès}". Une fois dans le dossier Info-hackaton, taper "code ." dans le terminal pour ouvrir VS Code directement dans ce dossier. Cela permet de garder les bons chemins d'accès pour les dossiers d'entrainement de l'IA

Ensuite, il faudra indiquer le chemin d'accès de la prise de note et du devis aux endroits dédiés, dans le programme docs.py. C'est indiqué dans la partie basse du programme. Il ne faut pas toucher au reste, car ce sont les documents qui permettent l'entraînement du modèle

Dans ce même programme, indiquer le lien sous lequel vous souhaitez recevoir le contrat généré

Finalement, il suffira de retourner dans le terminal, et d'écrire "python main.py". Cela lancera le programme pour générer le document. A noter que cette requête prendra une dizaine de minutes à s'effectuer, puisque la génération par IA s'effectue directement sur le CPU et GPU du pc, par soucis de confidentialité

Des warnings peuvent apparaître durant la génération, mais rien de bien grave. Cependant, en cas d'erreur, cherchez plutôt à contacter un responsable de maintenance informatique pour éviter d'insérer involontairement des erreurs dans le programme