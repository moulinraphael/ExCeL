//Affectation des variables globales
var socket = io.connect('http://localhost:8080');
var user;
var $active = false;
var $limitCells = 1000;


//Regex pour les différentes formules acceptées
var $regexCopy = /^=([A-Z]+[1-9][0-9]*) *$/;
var $regexInt = /^=INT\(([A-Z]+[1-9][0-9]*)\) *$/;
var $regexSup = /^=SUP\(([A-Z]+[1-9][0-9]*) *; *([A-Z]+[1-9][0-9]*)\) *$/;
var $regexSum = /^=SUM\((([A-Z]+[1-9][0-9]*):[A-Z]+[1-9][0-9]*)\) *$/;
var $regexAverage = /^=AVERAGE\((([A-Z]+[1-9][0-9]*):[A-Z]+[1-9][0-9]*)\) *$/;
var $regexMax = /^=MAX\((([A-Z]+[1-9][0-9]*):[A-Z]+[1-9][0-9]*)\) *$/;
var $regexCount = /^=COUNT\((([A-Z]+[1-9][0-9]*):[A-Z]+[1-9][0-9]*)\) *$/;
var $regexExpr = /^=/;


//Regex pour les expressions arithmétiques
var $regexRealNumber = "[-]?[0-9]*\\\.?[0-9]+";
var $regexCell = '[A-Z]+[1-9][0-9]*';
var $regexValue = '(?:' + $regexRealNumber + '|' + $regexCell + ')';
var $regexOperation = '[\*\-\+\/\^]';
var $regexTest = new RegExp('^=((' + $regexValue + $regexOperation + ')*(' + $regexValue + '))$');


//Lancement de l'application
$(function() {
    
    //Récupération du nom de l'utilisateur qui ne doit pas être vide
    while (!(user = prompt('Qui es-tu?'))) {}
    
    //Ajout des lignes et des colonnes
    $addRows(rows ? rows : ($('#table-excel').height() - 30) / 26 / 2, rows);
    $addColumns(columns ? columns : ($('#table-excel').width() - 50) / 100 / 2, rows);
    
    //Réinititialisation des ascenseurs
    $('#table-excel').scrollTop(0).scrollLeft(0);

    //Décalage des numéros de lignes et des identification de colonnes lors du scroll dans le tableau
    $('#table-excel').on('scroll', function() {
        $('#table-excel tr th:first-of-type div').css({'left':$('#table-excel').scrollLeft()});
        $('#table-excel thead tr th div').css({'top':$('#table-excel').scrollTop()});
    });

    //Emission des évènements initiaux
    socket.emit('adduser', id_file, user);
    socket.emit('load', id_file);
});


//Réception de l'évènement de création de connexion au serveur NodeJS
socket.on('connect', function (){
    socket.on('disconnected', function() {});
});


//Réception de l'évènement de chargement du contenu d'une cellule
socket.on('loadcell', function (cell) {
    if ($getCell(cell.cell))
        $getCell(cell.cell).val(cell.input).data('input', cell.input);
});


//Réception de l'évènement de chargement d'un style d'une cellule
socket.on('loadstyle', function (style) {
    if ($getCell(style.cell)) {
        $active = style.cell;
        if (style.group == 'style') $toggleStyle(style.style, true);
        else if (style.group == 'color') $toggleColor(style.style, true);
        else if (style.group == 'back') $toggleBack(style.style, true);
        else if (style.group == 'align') $toggleAlign(style.style, true);
        $active = false;
    }
});


//Réception de l'évènement de chargement de la feuille
//On calcule toutes les cellules 
socket.on('load', function () {
    $actionsOnBlur(null, null);
});


//Réception de l'évènement de mise à jour des personnes
socket.on('update', function (users){
    $('#users-excel').empty();
    for(var i in users) {
        $('#users-excel').append('<div title="'+users[i].user+'">' + users[i].user.substring(0,1).toUpperCase() + '</div>'); 
    }
});


//Réception de l'évènement permettant de voir la modification en temps réel d'une cellule
socket.on('onkey', function (socketid, cell, val){
    if (socketid != socket.id) {
        $getCell(cell).val(val);
        $getCell(cell).removeClass('cell-error').addClass('cell-edit');
    }
});


//Réception de l'évènement déclenché en sortie d'une sortie
//Modification effective et donc nécessité de recalculer les cellules
socket.on('onblur', function (socketid, cell, val){
    if (socketid != socket.id) {
        $getCell(cell).val(val).data('input', val).removeClass('cell-edit');
        $actionsOnBlur($getCell(cell), null);
    }
});


//Réception de l'évènement déclenché lors d'une modification de style
socket.on('toggle', function (socketid, cell, group, style){
    if (socketid != socket.id) {
        var $activeTemp = $active;
        $active = cell;
        if (group == 'style') $toggleStyle(style, true);
        if (group == 'color') $toggleColor(style, true);
        if (group == 'back') $toggleBack(style, true);
        if (group == 'align') $toggleAlign(style, true);
        $active = $activeTemp;
    }
});


//Traduction du numéro d'une colonne en lettre(s)
$colNumToLetter = function(n, fromZero) {
    var ordA = 'A'.charCodeAt(0);
    var ordZ = 'Z'.charCodeAt(0);
    var len = ordZ - ordA + 1;
  
    var s = "";

    if (!fromZero) 
        n = n - 1;

    while(n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s;
        n = Math.floor(n / len) - 1;
    }
    return s;
};


//Traduction du nom d'une colonne en numéro
$colLetterToNum = function(str, fromZero) {
    var ordA = 'A'.charCodeAt(0);
    var ordZ = 'Z'.charCodeAt(0);

    var p = 1;
    var r = 0;
    str = str.split("").reverse().join("");
    for (var i in str) {
        r = r + p * (str.charCodeAt(i) - ordA + 1);
        p = p * 26;
    }
    return r - (fromZero ? 1 : 0);
};


//Instanciation des actions sur les cellules
//On relance cette fonction lors de la création de nouvelles cellules
$actionsOnCells = function() {
    $('#table-excel tr td input')

    //Lorsque qu'une touche est appuyée
    .on('keydown', function(e) {
        $actionsOnKeyDown($(this), e);
    })

    //Lorsque qu'une touche a modifié le contenu
    .on('keyup', function(e) {
        $actionsOnKeyUp($(this), e);
    })

    //Lorsque l'utilisateur quitte une cellule
    .on('blur', function(e) {
        socket.emit('onblur', id_file, $(this).data('cell'), $(this).data('active') ? $(this).val() : $(this).data('input'));
        if ($(this).data('active')) 
            $(this).data('input', $(this).val());
        $actionsOnBlur($(this), e);
        $(this).data('active', false).removeClass('cell-active');
    })

    //Lorsque l'utilisateur arrive sur une cellule
    .on('focus', function(e) {
        $(this).val($(this).data('input'));
        $actionsOnFocus($(this), e);
    })

    //Lorsque l'utilisateur double clique sur une cellule
    .on('dblclick', function(e) {
        $actionsOnDblClick($(this), e);
    })

    //Par défaut, on affirme que la cellule n'est pas active
    .data('active', false);
}


//Actions réalisées lorsqu'une touche est appuyée
$actionsOnKeyDown = function(init, e) {
    var code = e.keyCode || e.which;
    var regex = /^([A-Z]+)([1-9][0-9]*)$/;
    var match = regex.exec(init.data('cell'));

    //Touche entrée
    if (code == 13) {
        if (init.data('active')) {
            init.data('input', init.val());
            socket.emit('onblur', id_file, init.data('cell'), init.val());
        }
        init.data('active', !init.data('active'));
    }

    //Touche bas
    else if (code == 40 && parseInt(match[2]) < parseInt($('#table-excel').data('rows'))) {
        $('input[data-cell="' + match[1] + (parseInt(match[2]) + 1) + '"]').focus();
        e.preventDefault();
    }

    //Touche droite
    else if ((!init.data('active') || init[0].selectionStart == init.val().length) && 
        code == 39 && $colLetterToNum(match[1]) < parseInt($('#table-excel').data('columns'))) {
        $('input[data-cell="' + $colNumToLetter($colLetterToNum(match[1]) + 1) + match[2] + '"]').focus();
        e.preventDefault();
    }

    //Touche gauche
    else if ((!init.data('active') || init[0].selectionStart == 0) && code == 37 && match[1] != 'A') {
        $('input[data-cell="' + $colNumToLetter($colLetterToNum(match[1]) - 1) + match[2] + '"]').focus();
        e.preventDefault();
    }

    //Touche haut
    else if (code == 38 && match[2] != '1') {
        $('input[data-cell="' + match[1] + (parseInt(match[2]) - 1) + '"]').focus();
        e.preventDefault();
    }

    //Pour toutes les autres touches, on affirme que la cellule est active
    else if ($.inArray(code, [13, 26, 37, 38, 39, 40]) < 0)
        init.data('active', true);

    //Dans le cas où la cellule est active, on change le style et 
    //on émet un signal pour modifier le contenu chez tous les utilisateurs
    if (init.data('active')) {
        socket.emit('onkey', id_file, init.data('cell'), init.val());
        init.addClass('cell-active');
    }

    //Sinon on supprimer le style de cellule active
    else
        init.removeClass('cell-active');
};


//Actions lors du double clique, identique à la touche entrée
$actionsOnDblClick = function(init, e) {
    init.data('active', true).addClass('cell-active');
};


//Actions lors de la modification du contenu
//Emission d'un signal
$actionsOnKeyUp = function(init, e) {
    if (init.data('active'))
        socket.emit('onkey', id_file, init.data('cell'), init.val());
};


//Actions lors du focus d'une cellule, on modifie le style
$actionsOnFocus = function(init, e) {
    $active = init.data('cell');
    init.removeClass('cell-error');
};


//Actions lorsque l'utilisateur quitte une cellule
//On recalcule toutes les cellules  
$actionsOnBlur = function(init, e) {
    $recalcCell(init);
    $('#table-excel tbody tr td input').removeData('actualise').removeData('error').removeClass('cell-error').each(function() {
        $recalcCell($(this));
    });
};


//Fonction permettant de calculer la valeur d'une cellule en fonction du contenu pouvant être une formule
$recalcCell = function(cell, init) {
    if (!cell)
        return;

    var match;
    var inits = (!init ? cell.data('cell') : init + ',' + cell.data('cell')) + ',';

    //La cellule est déjà en erreur
    if (cell.data('error')) {
        cell.addClass('cell-error');
    }

    //Une boucle est trouvée
    else if (init && init.indexOf(cell.data('cell')+',') >= 0) {
        cell.val('#Loop#').data('error', true).addClass('cell-error');
    }

    //La cellule est vide et mal initialisée
    else if (!cell.data('actualise') && (
            typeof(cell.data('input')) == 'undefined' ||
            cell.data('input') == '')) {
        cell.data('input', '');
    }

    //Copie d'une cellule dans une autre =XN
    else if ((match = $regexCopy.exec(cell.data('input'))) !== null) {
        
        //La cellule exite
        if ($getCell(match[1])) {
            
            //On recalcule la cellule référencée
            $recalcCell($getCell(match[1]), inits);
            
            //La cellule référencée est la même, c'est une boucle
            if (cell.data('cell') == match[1])
                cell.data('error', true).addClass('cell-error').val('#Loop#');
            
            //Il y a une erreur dans la cellule référencée
            else if ($getCell(match[1]).data('error'))
                cell.data('error', true).addClass('cell-error').val('#ValRef#');
            
            //La cellule référencée est vide, c'est plus une information qu'une erreur
            else if ($getCell(match[1]).val() == '')
                cell.data('error', true).addClass('cell-error').val('#Empty#');
            
            //Sinon on copie la valeur
            else
                cell.val($getCell(match[1]).val());
        }

        //La cellule n'exite pas
        else
            cell.data('error', true).addClass('cell-error').val('#Out#');
    }

    //Formule qui retourne la partie entière d'un nombre
    else if ((match = $regexInt.exec(cell.data('input'))) !== null) {
        
        //La cellule exite
        if ($getCell(match[1])) {

            //On recalcule la cellule référencée
            $recalcCell($getCell(match[1]), inits);
            
            //La cellule référencée est la même, c'est une boucle
            if (cell.data('cell') == match[1])
                cell.data('error', true).addClass('cell-error').val('#Loop#');
            
            //Il y a une erreur dans la cellule référencée
            else if ($getCell(match[1]).data('error'))
                cell.data('error', true).addClass('cell-error').val('#ValRef#');
            else if (!$getCell(match[1]).val())
                cell.val(0);
            else if (!$.isNumeric($getCell(match[1]).val()))
                cell.data('error', true).addClass('cell-error').val('#NaN#');
            else
                cell.val(parseInt($getCell(match[1]).val()));
        }

        //La cellule n'exite pas
        else
            cell.data('error', true).addClass('cell-error').val('#Out#');
    }

    //Formule qui regarder si la valeur d'une cellule est supérieur à la valeur d'une autre
    else if ((match = $regexSup.exec(cell.data('input'))) !== null) {
        
        //Les deux cellules existent
        if ($getCell(match[1]) && $getCell(match[2])) {
            
            //On recalcule les deux cellules
            $recalcCell($getCell(match[1]), inits);
            $recalcCell($getCell(match[2]), inits);
            
            //Si l'une des cellules référencées est l'hôte alors il y a une boucle
            if (cell.data('cell') == match[1] ||
                cell.data('cell') == match[2])
                cell.data('error', true).addClass('cell-error').val('#Loop#');
            
            //L'une des cellules contient une erreur
            else if ($getCell(match[1]).data('error') ||
                $getCell(match[2]).data('error'))
                cell.data('error', true).addClass('cell-error').val('#ValRef#');
            
            //L'une des cellules n'a pas de valeur numérique
            else if (!$.isNumeric($getCell(match[1]).val()) ||
                !$.isNumeric($getCell(match[2]).val()))
                cell.data('error', true).addClass('cell-error').val('#NaN#');
            
            //Sinon on compare les deux valeurs 
            else 
                cell.val($getCell(match[1]).val() > $getCell(match[2]).val() ? '@True' : '@False');
            
        }

        //L'une des deux cellules au moins n'existe pas
        else 
            cell.data('error', true).addClass('cell-error').val('#Out#');
    } 

    //Evaluation d'une expression arithmétique sans parenthèse
    else if ((match = $regexTest.exec(cell.data('input'))) !== null) {
        var regex = new RegExp($regexCell, 'g');
        var error = false;
        var last = match[3];
        
        //Récupération de chacune des cellules référencées
        while (cells = regex.exec(cell.data('input'))) {
            
            //S'il y a déjà une erreur rencontrée on quitte la boucle
            if (error)
                break;

            //La cellule n'exite pas
            else if (!$getCell(cells)) {
                cell.data('error', true).addClass('cell-error').val('#Out#');
                error = true;
            }

            //La cellule référencée est la cellule hôte, c'est une boucle
            else if (cell.data('cell') == cells) {
                cell.data('error', true).addClass('cell-error').val('#Loop#');
                error = true;
            }

            //Il y a déjà une erreur dans la cellule référencée
            else if ($getCell(cells).data('error')) {
                cell.data('error', true).addClass('cell-error').val('#ValRef#');
                error = true;
            }

            //La cellule référencée n'est pas une valeur numérique
            else if (!$.isNumeric($getCell(cells).val())) {
                cell.data('error', true).addClass('cell-error').val('#NaN#');
                error = true;
            }

            //Sinon tout va bien, on recalcule la cellule
            else
                $recalcCell($getCell(cells));
        }

        //Pas d'erreur, on peut remplacer les cellules par leur valeur puis calculer l'expression
        if (!error) {
            $regexOperandes = new RegExp('^(' + $regexCell + '|' + $regexRealNumber + ')(' + $regexOperation + ')');
            var matches;
            var expr = '';
            var it = 0;
            var result = null;

            //On remplace les cellules par leur valeur numérique
            //On évite une boucle infinie au cas où
            while (it++ < 1000 && (matches = $regexOperandes.exec(match[1]))) {
                expr = expr + ($getCell(matches[1]) !== undefined ? $getCell(matches[1]).val() : matches[1]) + matches[2];
                match[1] = match[1].substring(matches[0].length);
            }

            //De même pour le dernier élement si c'est une cellule
            expr = expr + ($getCell(last) !== undefined ? $getCell(last).val() : last);
            
            //On évalue l'expression à l'aide de la bibliothèque Parser
            result = Parser.evaluate(expr);
            cell.val(Math.round(result * 1000) / 1000);
        }
    }

    //Expression pour calculer la somme d'une plage de données
    else if ((match = $regexSum.exec(cell.data('input'))) !== null) {
        
        //La plage n'est pas valide
        if (!$isRangeValid(match[1]))
            cell.data('error', true).addClass('cell-error').val('#Range#');
        
        //La plage est valide
        else {
            var $somme = 0;
            var $elem = match[2];
            
            //Pour chaque cellule de la plage en commencant par la première
            do {

                //La cellule n'existe pas
                if (!$getCell($elem)) {
                    $somme = '#Out#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On recalcule la cellule
                $recalcCell($getCell($elem), inits);
                
                //La cellule étudiée de la plage est la cellule hôte, c'est une boucle
                if ($elem == cell.data('cell')) {
                    $somme = '#Loop#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //La cellule étudiée de la plage ne possède pas de valeur numérique
                else if (!$.isNumeric($getCell($elem).val())) {
                    $somme = '#NaN#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On somme
                $somme += parseFloat($getCell($elem).val());
            } while ($elem = $getNextInRange(match[1], $elem));
            
            //On affiche le résultat que ce soit une erreur ou non
            cell.val($.isNumeric($somme) ? Math.round($somme * 1000) / 1000 : $somme);
        }
    } 

    //Expression pour calculer la moyenne d'une plage
    else if ((match = $regexAverage.exec(cell.data('input'))) !== null) {
        
        //La plage n'est pas valide
        if (!$isRangeValid(match[1]))
            cell.data('error', true).addClass('cell-error').val('#Range#');
        
        //La plage est valide
        else {
            var $somme = 0;
            var $quantite =0;
            var $elem = match[2];
            
            //Pour chaque cellule de la plage en commencant par la première
            do {
                
                //La cellule n'existe pas
                if (!$getCell($elem)) {
                    $somme = '#Out#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On recalcule la cellule
                $recalcCell($getCell($elem), inits);
                
                //La cellule étudiée de la plage est la cellule hôte, c'est une boucle
                if ($elem == cell.data('cell')) {
                    $somme = '#Loop#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //La cellule étudiée de la plage ne possède pas de valeur numérique
                else if (!$.isNumeric($getCell($elem).val())) {
                    $somme = '#NaN#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On somme et on incrémente le nombre de cellules
                $somme += parseFloat($getCell($elem).val());
                $quantite += 1;
            } while ($elem = $getNextInRange(match[1], $elem));
            
            //On affiche l'erreur s'il y en a une
            if(!$.isNumeric($somme))
                cell.val($somme);
            
            //On affiche le résultat
            else
                cell.val(Math.round($somme / $quantite * 1000) / 1000);
        }
    }

    //Expression permettant de trouver la valeur max d'une plage
    else if ((match = $regexMax.exec(cell.data('input'))) !== null) {
        
        //La plage n'est pas valide
        if (!$isRangeValid(match[1]))
            cell.data('error', true).addClass('cell-error').val('#Range#');
        
        //La plage est valide
        else {
            var $max;
            var $elem = match[2];
            
            //Pour chaque cellule de la plage en commencant par la première
            do {
                
                //La cellule n'existe pas
                if (!$getCell($elem)) {
                    $max = '#Out#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On recalcule la cellule
                $recalcCell($getCell($elem), inits);
                
                //La cellule étudiée de la plage est la cellule hôte, c'est une boucle
                if ($elem == cell.data('cell')) {
                    $max = '#Loop#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //La cellule ne possède pas de valeur numérique
                else if (!$.isNumeric($getCell($elem).val())) {
                    $max = '#NaN#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //Comparaison avec le maximum temporaire
                else if(typeof($max) == 'undefined' || 
                    parseFloat($getCell($elem).val()) > $max){
                    $max = parseFloat($getCell($elem).val());
                }
            } while ($elem = $getNextInRange(match[1], $elem));
            
            //On affiche le résultat ou l'erreur
            cell.val($max);
        }
    }

    //Expression pour compter le nombre de cellules non vides
    else if ((match = $regexCount.exec(cell.data('input'))) !== null) {
        
        //La plage n'est pas valide
        if (!$isRangeValid(match[1]))
            cell.data('error', true).addClass('cell-error').val('#Range#');
        
        //La plage est valide
        else {
            var $count = 0;
            var $elem = match[2];
            
            //Pour chaque cellule de la plage en commencant par la première
            do {
                
                //La cellule n'existe pas
                if (!$getCell($elem)) {
                    $count = '#Out#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //On recalcule la cellule
                $recalcCell($getCell($elem), inits);
                
                //La cellule étudiée de la plage est la cellule hôte, c'est une boucle
                if ($elem == cell.data('cell')) {
                    $count = '#Loop#';
                    cell.data('error', true).addClass('cell-error');
                    break;
                }

                //Sinon on incrémente
                else if($getCell($elem).val() != ''){
                    $count++;
                }
            } while ($elem = $getNextInRange(match[1], $elem));
            
            //On affiche le résultat ou l'erreur
            cell.val($count);
        }
    }

    //Toutes les autres expressions non reconnues
    else if ($regexExpr.test(cell.data('input'))) 
        cell.val('#Expr#').data('error', true).addClass('cell-error');
    

    //La cellule est mise en tant que cellule actualisée
    cell.data('actualise', true);
};


//Ajout d'une colonne
$addColumn = function() {
    var $columns = parseInt($('#table-excel').data('columns'));
    var $rows = parseInt($('#table-excel').data('rows'));

    //Si le nombre de cellule est trop important, on demande l'accord
    if ($columns * $rows > $limitCells && !confirm("Rajouter des cellules ralenti la page.\nVoulez-vous continuer?"))
        return;

    //Pour chaque ligne on ajoute un élément à la fin
    $('#table-excel tr').each(function(n) {
        
        //Si c'est le premier on décale le lien pour l'ajout de colonne
        if (n == 0)
            $(this).find('th:last').html('<div>' + $colNumToLetter($columns) + '</div>').end()
                .append('<th><div>' + $colNumToLetter($columns + 1) +
                    '<a class="add-column" onclick="$addColumns(1)"></a></div></th>');
        else
            $(this).append('<td><input type="text" data-cell="' + $colNumToLetter($columns + 1) + n + '" /></td>');
    });

    $('#table-excel').data('columns', $columns + 1);
};


//Ajout d'une ligne
$addRow = function() {
    var $columns = parseInt($('#table-excel').data('columns'));
    var $rows = parseInt($('#table-excel').data('rows'));

    //Si le nombre de cellule est trop important, on demande l'accord
    if ($columns * $rows > $limitCells && !confirm("Rajouter des cellules ralenti la page.\nVoulez-vous continuer?"))
        return;

    //Décalage du lien pour l'ajout de ligne
    $('#table-excel').find('tr:last th:first').html('<div>' + $rows + '</div>').end().append('<tr><th><div>' + ($rows + 1) + 
        '<a class="add-row" onclick="$addRows(1)"></a></div></th></tr>');

    //Pour chaque colonne on rajoute une cellule
    for (var i = 1; i < $("#table-excel tr:first > th").length; i++ )
        $('#table-excel tr:last').append('<td><input type="text" data-cell="' + $colNumToLetter(i) + ($rows + 1) + '" /></td>');
    
    $('#table-excel').data('rows', $rows + 1);
};


//Ajout de plusieurs lignes et mise à jour des actions
$addRows = function(n, loop) {
    for (var j = 0; j < n; j++)
        $addRow();

    //Mise à jour des actions, recalcul des cellules et décalage de la page
    $actionsOnCells();
    $actionsOnBlur(null, null);
    $('#table-excel').scrollTop($('#table-excel').innerHeight());
    
    //Lorsque c'est une action de l'utilisateur, on émet un évènement
    if (!loop)
        socket.emit('majrows', id_file, $('#table-excel').data('rows'));
};


//Ajout de plusieurs colonnes et mise à jour des actions
$addColumns = function(n, loop) {
    for (var j = 0; j < n; j++)
        $addColumn();

    //Mise à jour des actions, recalcul des cellules et décalage de la page
    $actionsOnCells();
    $actionsOnBlur(null, null);
    $('#table-excel').scrollLeft($('#table-excel').innerWidth());
    
    //Lorsque c'est une action de l'utilisateur, on émet un évènement
    if (!loop)
        socket.emit('majcolumns', id_file, $('#table-excel').data('columns'));
};


//Validation d'une plage de cellules
$isRangeValid = function(str) {
    var regex = /^([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)$/;
    var match = regex.exec(str);

    //Invalide
    if (match === null ||
        parseInt(match[2]) > parseInt(match[4]) ||
        $colLetterToNum(match[1]) > $colLetterToNum(match[3]))
        return false;

    return true;
};


//Récupération de la cellule suivante dans une plage
$getNextInRange = function(range, cell) {
    var regex_range = /^([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)$/;
    var regex_cell = /^([A-Z]+)([1-9][0-9]*)$/;
    var match_range = regex_range.exec(range);
    var match_cell = regex_cell.exec(cell);

    //Dernière cellule
    if (match_cell[1] == match_range[3] && match_cell[2] == match_range[4])
        return false;

    var row = match_cell[1] == match_range[3] ? parseInt(match_cell[2]) + 1 : match_cell[2];
    var column = match_cell[1] == match_range[3] ? match_range[1] : 
        $colNumToLetter($colLetterToNum(match_cell[1]) + 1);
    return column + row; 
};


//Récupération de la cellule (objet) à partir de son identification
$getCell = function(str) {
    var regex = /([A-Z]+)([1-9][0-9]*)/;
    var $columns = parseInt($('#table-excel').data('columns'));
    var $rows = parseInt($('#table-excel').data('rows'));

    var match = regex.exec(str);
    if (match === null)
        return;

    var column = $colLetterToNum(match[1]);
    var row = parseInt(match[2]);
    if (column > $columns || row > $rows)
        return;

    return $('#table-excel tbody tr:nth(' + (row - 1) + ') td:nth(' + (column - 1) + ') input');
};


//Changement du style de format d'une cellule
$toggleStyle = function(style, loop) {
    
    //Si aucune cellule n'est active, ça ne sert à rien de continuer
    if (!$active)
        return;

    //Action utilisateur, on émet un signal
    if (!loop)
        socket.emit('toggle', id_file, $active, 'style', style, $getCell($active).hasClass('cell-style-'+style));

    //Mise à jour du style
    $getCell($active).toggleClass('cell-style-'+style);

    //On revient sur la cellule lors d'une action utilisateur
    if (!loop)
        $getCell($active).focus();
};


//Changement de la couleur de texte d'une cellule
$toggleColor = function(color, loop) {
    
    //On cache le bloc des couleurs
    if (!loop)
        $('#tools-excel-color').toggle();

    //Si aucune cellule n'est active, ça ne sert à rien de continuer
    if (!$active)
        return;

    var has = $getCell($active).hasClass('cell-color-' + color);

    //Action utilisateur, on émet un signal
    if (!loop)
        socket.emit('toggle', id_file, $active, 'color', color, has);
    
    //On supprime toutes les classes liées à la couleur de texte
    if ($getCell($active).attr('class'))
        $getCell($active).attr('class', $getCell($active).attr('class').replace(/\bcell\-color.*?\b/g, ''));
    
    //S'il n'a pas la couleur on l'ajoute
    if (!has)
        $getCell($active).addClass('cell-color-' + color);

    //On revient sur la cellule lors d'une action utilisateur
    if (!loop)
        $getCell($active).focus();
};

$toggleBack = function(color, loop) {
    
    //On cache le bloc des couleurs de fond
    if (!loop)
        $('#tools-excel-back').toggle();

    //Si aucune cellule n'est active, ça ne sert à rien de continuer
    if (!$active)
        return;

    var has = $getCell($active).hasClass('cell-back-' + color);

    //Action utilisateur, on émet un signal
    if (!loop)
        socket.emit('toggle', id_file, $active, 'back', color, has);
    
    //On supprime toutes les classes liées à la couleur de texte
    if ($getCell($active).attr('class'))
        $getCell($active).attr('class', $getCell($active).attr('class').replace(/\bcell\-back.*?\b/g, ''));
    
    //S'il n'a pas le fond on l'ajoute
    if (!has)
        $getCell($active).addClass('cell-back-' + color);

    //On revient sur la cellule lors d'une action utilisateur
    if (!loop)
        $getCell($active).focus();
};

$toggleAlign = function(align, loop) {
    
    //Si aucune cellule n'est active, ça ne sert à rien de continuer
    if (!$active)
        return;

    var has = $getCell($active).hasClass('cell-align-' + align);

    //Action utilisateur, on émet un signal
    if (!loop)
        socket.emit('toggle', id_file, $active, 'align', align, has);
    
    //On supprime toutes les classes liées à l'alignement'  
    if ($getCell($active).attr('class'))
        $getCell($active).attr('class', $getCell($active).attr('class').replace(/\bcell\-align.*?\b/g, ''));
    
    //S'il n'a pas l'alignement on l'ajoute
    if (!has)
        $getCell($active).addClass('cell-align-' + align);

    //On revient sur la cellule lors d'une action utilisateur
    if (!loop)
        $getCell($active).focus();
};