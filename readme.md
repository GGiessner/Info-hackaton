Le projet consiste en la rédaction d'une proposition commerciale pour un projet de conseil en stratégie.

Il s'agit de réaliser de manière complètement automatique un contrat, en utilisant des notes prises pendant l'entretien, ainsi qu'un devis fait afin de répondre à la problématique du client. Nous avons également à notre disposition un Excel qui contient pleins de références d'anciennes missions, afin de construire un liste de références de missions réalisées similaires à la mission actuelle.
L'entreprise souhaite que nous utilisions des LLM afin de pousser aussi loin que possible l'automatisation de ce projet.


Pour effectuer cette mission proposée par l'entreprise GJOA, nous avons décidé d'écouter leurs conseils et d'utiliser des LLM.
Dans un premier temps, nous venons chercher les informations utiles au projet et diirectement utilisables telles qu'elles sont (le nom de la société, la date...).
Ensuite, nous utilisons les LLM pour traiter les notes prises pendant l'entretien. Cela nous permet de générer la contextualisation de la mission dans le contrat. Cela va également nous permettre de trouver des références de missions proches de ce que nous pouvons trouver dans le Excel des références similaires au projet en cours. Finalement, à partir de requêtes dans le devis, nous pouvons extraire les différents axes de travail du contrat, et nous pouvons donc, à partir d'une succinte description, générer un petit texte pour décrire les différentes sous-parties du programme.

Afin d'effectuer le contrat dans la plus grande confidentialité, nous faisons tourner le programme sur un programme de LLM en local. Cela évite toute potentielle fite d'information, mais rend le traitement plus long et potentiellement coûteux en termes de Hardware.

Pour ce projet, nous utilisons une librairie un peu exotique: Ollama