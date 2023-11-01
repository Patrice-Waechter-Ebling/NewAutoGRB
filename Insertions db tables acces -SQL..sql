insert into [WebGRB2023].[dbo].GrbAchat select * from [GRB2023].[dbo].GRB_Achat
insert into [WebGRB2023].[dbo].GrbAchatPiece select IDAchat, IndexAchat, PIECE, NuméroLigne, Qté, Desc_FR, Desc_EN, Manufact, Prix_list, Escompte, Prix_net, IDFRS, Prix_total, Type, Commandé, Retour, NoRetour, Recu, DateRéception, QuantitéRecue, DateCommande, DateRequise, Inutile, CommandeAnnulée, DateRetour, PrixOrigine, Devise from [GRB2023].[dbo].GRB_Achat_Pieces
insert into [WebGRB2023].[dbo].GrbAchatPiecesTampon select DateCopie, Initiales, IDAchat, IndexAchat, PIECE, NuméroLigne, Qté, Desc_FR, Desc_EN, Manufact, Prix_list, Escompte, Prix_net, IDFRS, Prix_total, Type, Commandé, Retour, NoRetour, Recu, DateRéception, QuantitéRecue, DateCommande, DateRequise, Inutile, CommandeAnnulée, DateRetour, PrixOrigine, Devise from [GRB2023].[dbo].GRB_Achat_Pieces_Tampon
insert into [WebGRB2023].[dbo].GrbAlarme select Initiale, IDCédule, Type, Date, Heure, Message, JourSemaine, NoEmploye, TypeCédule from [GRB2023].[dbo].GRB_Alarmes
insert into [WebGRB2023].[dbo].GrbAutorisationPunch select NoEmploye, AutoriserPar from [GRB2023].[dbo].GRB_AutorisationPunch
insert into [WebGRB2023].[dbo].GrbBavardSuppression select  IDUser, NoProjSoum, Type, Qté, [No Item], Date, Heure from [GRB2023].[dbo].GRB_BavardSuppression
insert into [WebGRB2023].[dbo].GrbBonsCommande select * from [GRB2023].[dbo].GRB_BonsCommandes
insert into [WebGRB2023].[dbo].GrbBonsCommandesPiece select * from [GRB2023].[dbo].GRB_BonsCommandes_Pieces
insert into [WebGRB2023].[dbo].GrbCatalogueElec select * from [GRB2023].[dbo].GRB_CatalogueElec
insert into [WebGRB2023].[dbo].GrbCatalogueMec select * from [GRB2023].[dbo].GRB_CatalogueMec
insert into [WebGRB2023].[dbo].GrbCedule select  initiale, date_cedulé, heure_début, heure_fin, client, joursemaine, transport, finprojet, Alarme from [GRB2023].[dbo].GRB_cédule
insert into [WebGRB2023].[dbo].GrbClient select NomClient, Compagnie, Telephonne, Fax, Pagette, Cellulaire, ContactGRB, Email, AdresseLiv, VilleLiv, [Prov/EtatLiv], PaysLiv, CodePostalLiv, noposte, Commentaire, SiteWeb, DateCréation, UserCréation, DateModification, UserModification, EntryIDOutlook, Béton, Pavé, Pharmaceutique, Agroalimentaire, Meuble, Meunerie, Manufacturier, Autre, Consultant, Asphalte, ICPI, Potentiel, ProduitsChimiques, Supprimé, NomContact from [GRB2023].[dbo].GRB_Client
insert into [WebGRB2023].[dbo].GrbCommentaire select NoProjSoum, [Index], Commentaire, Section, SousSection, [Key], [Relative] from [GRB2023].[dbo].GRB_Commentaires
insert into [WebGRB2023].[dbo].GrbConfig select * from [GRB2023].[dbo].GRB_Config
insert into [WebGRB2023].[dbo].GrbContact select NomContact, Compagnie, Telephonne, Fax, Cellulaire, Pagette, [E-mail], noposte, teldomicile, DateCréation, UserCréation, DateModification, UserModification, Commentaire, Titre, EntryIDOutlook, Supprimé from [GRB2023].[dbo].GRB_contact
insert into [WebGRB2023].[dbo].GrbContactClient select  noclient, nocontact from [GRB2023].[dbo].GRB_ContactClient
insert into [WebGRB2023].[dbo].GrbContactFr select NoFRS, NoContact from [GRB2023].[dbo].GRB_ContactFRS
insert into [WebGRB2023].[dbo].GrbDessin select  NoProjet, NoDessin, Description, Type from [GRB2023].[dbo].GRB_Dessins
insert into [WebGRB2023].[dbo].GrbDoublon select  PIECE, FABRICANT, DESCR_FR, DESCR_EN, CATEGORIE from [GRB2023].[dbo].GRB_Doublons
insert into [WebGRB2023].[dbo].GrbEmploye select * from [GRB2023].[dbo].GRB_employés
insert into [WebGRB2023].[dbo].GrbErreur select Qui, Date, Heure, Form, Methode, NoLigne, NoErreur, Description, Source, Params from [GRB2023].[dbo].GRB_Erreurs
insert into [WebGRB2023].[dbo].GrbExceptionsDl select Exception from [GRB2023].[dbo].GRB_ExceptionsDL
insert into [WebGRB2023].[dbo].GrbFamille select [Famille] from [GRB2023].[dbo].GRB_Famille
insert into [WebGRB2023].[dbo].GrbFournisseur select * from [GRB2023].[dbo].GRB_Fournisseur
insert into [WebGRB2023].[dbo].GrbGroupe select NomGroupe, Clients, Fournisseurs, Contacts, ContactsVendeurs, Rapport, CatalogueMec, CatalogueElec, Employes, Cedule, Configuration, Punch, Outils, SoumissionMec, ProjetMec, SoumissionElec, ProjetElec, InventaireMec, InventaireElec, Achat, ModificationFacturation, ModificationClients, ModificationFournisseurs, ModificationContacts, ModificationGroupes, ModificationEmployes, ModificationFeuillesTemps, ModificationOutils, ModificationSoumissionsMec, ModificationProjetsMec, ModificationSoumissionsElec, ModificationProjetsElec, ModificationBonsCommandes, ModificationCatalogueElec, ModificationCatalogueMec, ModificationInventaireMec, ModificationInventaireElec, ModificationPunchEmployes, ModificationReception, ModificationRetourMarchandise, SuppressionProjets, ListeDistribution, PunchSemaineAntérieure, VerrouillageTempsProjet, DéverrouillageTempsProjet from [GRB2023].[dbo].GRB_Groupes
insert into [WebGRB2023].[dbo].GrbImpressionBonlivraison select qte_com, qte_livr, qte_bo, description, manufacturier, user from [GRB2023].[dbo].GRB_impression_bonlivraison
insert into [WebGRB2023].[dbo].GrbImpressionDemandePrixElec select * from [GRB2023].[dbo].GRB_ImpressionDemandePrixElec
insert into [WebGRB2023].[dbo].GrbImpressionDemandePrixMec select * from [GRB2023].[dbo].GRB_ImpressionDemandePrixMec
insert into [WebGRB2023].[dbo].GrbImpressionDetailTemp select Employe, Type, TotalHeures from [GRB2023].[dbo].GRB_ImpressionDetailTemps
insert into [WebGRB2023].[dbo].GrbImpressionListePiece select  IDSoumission, nomSection, SousSection, NumItem, Qté, Description, Manufact, Section, NomSousSection, IDSection, ID from [GRB2023].[dbo].GRB_impression_ListePiece
insert into [WebGRB2023].[dbo].GrbImpressionPoste select * from [GRB2023].[dbo].GRB_ImpressionPoste
insert into [WebGRB2023].[dbo].GrbImpressionPunch select * from [GRB2023].[dbo].GRB_ImpressionPunch
insert into [WebGRB2023].[dbo].GrbImpressionSommairePunchGeneral select NoProjet, Total from [GRB2023].[dbo].GRB_ImpressionSommairePunchGeneral
insert into [WebGRB2023].[dbo].GrbImpressionSommairePunchGrb select Employé, NoProjet, Date, Commentaire, HeureDébut, HeureFin, NbreKM, Total from [GRB2023].[dbo].GRB_ImpressionSommairePunchGRB
insert into [WebGRB2023].[dbo].GrbImpressionSoumission select IDSoumission, nomSection, NumItem, Qté, Description, Manufact, Prix_list, Escompte, Prix_net, NomFournisseur, Temps, Temps_total, Prix_total, Profit_Pourcent, Profit_Argent, SousSection, DateReception, DateCommande, NoSéquentiel from [GRB2023].[dbo].GRB_impression_soumission
insert into [WebGRB2023].[dbo].GrbInventaireElec select  NoItem, Description, Manufacturier, QteBoite, [Prix Liste], Escompte, [Prix net], Commentaires, Localisation, DeviseMonétaire, QuantitéStock, QuantitéCommandée, Minimum, QuantitéMinimum, Commande, NoProjet, CommandeParBoite from [GRB2023].[dbo].GRB_InventaireElec
insert into [WebGRB2023].[dbo].GrbInventaireElecModif select  Date, IDProjet, NoItem, Quantité, User from [GRB2023].[dbo].GRB_InventaireElecModif
insert into [WebGRB2023].[dbo].GrbInventaireMec select  NoItem, Description, Manufacturier, QteBoite, [Prix Liste], Escompte, [Prix net], Commentaires, Localisation, DeviseMonétaire, QuantitéStock, QuantitéCommandée, Minimum, QuantitéMinimum, Commande, NoProjet, CommandeParBoite from [GRB2023].[dbo].GRB_InventaireMec
insert into [WebGRB2023].[dbo].GrbInventaireMecModif select  Date, IDProjet, NoItem, Quantité, User from [GRB2023].[dbo].GRB_InventaireMecModif
insert into [WebGRB2023].[dbo].GrbOutil select * from [GRB2023].[dbo].GRB_Outils
insert into [WebGRB2023].[dbo].GrbOutilsInOut select * from [GRB2023].[dbo].GRB_Outils_In_out
insert into [WebGRB2023].[dbo].GrbPiecesFr select * from [GRB2023].[dbo].GRB_PiecesFRS
insert into [WebGRB2023].[dbo].GrbProjetsDessins select * from [GRB2023].[dbo].GRB_ProjetsDessins
insert into [WebGRB2023].[dbo].GrbProjetElec select * from [GRB2023].[dbo].GRB_ProjetElec
insert into [WebGRB2023].[dbo].GrbProjetMec select * from [GRB2023].[dbo].GRB_ProjetMec
insert into [WebGRB2023].[dbo].GrbProjetModif select * from [GRB2023].[dbo].GRB_Projet_Modif
insert into [WebGRB2023].[dbo].GrbProjetPiece select * from [GRB2023].[dbo].GRB_Projet_Pieces
insert into [WebGRB2023].[dbo].GrbProjetPiecesTampon select * from [GRB2023].[dbo].GRB_Projet_Pieces_Tampon
insert into [WebGRB2023].[dbo].GrbProjSoum select * from [GRB2023].[dbo].GRB_ProjSoum
insert into [WebGRB2023].[dbo].GrbPunch select * from [GRB2023].[dbo].GRB_Punch
insert into [WebGRB2023].[dbo].GrbPunchExcel select * from [GRB2023].[dbo].GRB_PunchExcel
insert into [WebGRB2023].[dbo].GrbSortieMateriel select * from [GRB2023].[dbo].GRB_SortieMatériel
insert into [WebGRB2023].[dbo].GrbSoumissionElec select * from [GRB2023].[dbo].GRB_SoumissionElec
insert into [WebGRB2023].[dbo].GrbSoumissionMec select * from [GRB2023].[dbo].GRB_SoumissionMec
insert into [WebGRB2023].[dbo].GrbSoumissionModif select * from [GRB2023].[dbo].GRB_Soumission_Modif
insert into [WebGRB2023].[dbo].GrbSoumissionPiece select * from [GRB2023].[dbo].GRB_Soumission_Pieces
insert into [WebGRB2023].[dbo].GrbSoumissionPiecesTampon select * from [GRB2023].[dbo].GRB_Soumission_Pieces_Tampon
insert into [WebGRB2023].[dbo].GrbSoumProjSectionElec select * from [GRB2023].[dbo].GRB_SoumProjSectionElec
insert into [WebGRB2023].[dbo].GrbSoumProjSectionMec select * from [GRB2023].[dbo].GRB_SoumProjSectionMec
insert into [WebGRB2023].[dbo].GrbTempDp select * from [GRB2023].[dbo].GRB_TempDP
insert into [WebGRB2023].[dbo].GrbTempInventaire select * from [GRB2023].[dbo].GRB_TempInventaire
insert into [WebGRB2023].[dbo].GrbTransport select * from [GRB2023].[dbo].GRB_Transport
insert into [WebGRB2023].[dbo].GrbVendeur select * from [GRB2023].[dbo].GRB_vendeur
insert into [WebGRB2023].[dbo].ProjetPieceBack select * from [GRB2023].[dbo].Projet_piece_back
insert into [WebGRB2023].[dbo].ProjetTamponBack select * from [GRB2023].[dbo].projet_tampon_back
insert into [WebGRB2023].[dbo].TableDesErreur select * from [GRB2023].[dbo].[Table des erreurs]
insert into [WebGRB2023].[dbo].TblCategorie select * from [GRB2023].[dbo].TBL_Categorie
insert into [WebGRB2023].[dbo].TblPunchType select * from [GRB2023].[dbo].TBL_Punch_Type
