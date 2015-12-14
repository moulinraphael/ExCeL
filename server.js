var users = [];

var app = require('express')();
var mysql = require('mysql');
var server = require('http').Server(app);
var io = require('socket.io')(server);
var connection = mysql.createConnection({
    host     : 'localhost',
    user     : 'root',
    password : '1514co79m',
    database : 'nodejs_excel'
});


//On utilise le moteur de template EJS
app.engine('html', require('ejs').renderFile);
server.listen(8080);


//Page de listing des feuilles
app.get('/', function (req, res){
    res.sendFile(__dirname + '/listing.html');
});


//Page d'édition d'une feuille
app.get('/:id(\\d+)/', function (req, res){
    
    //On récupère les données propre à la feuille
    connection.query('SELECT * FROM files WHERE id = '+req.params.id, function(err, rows, fields) {        
        
        //La feuille n'existe pas
        if (!rows || !rows.length)
            res.sendFile(__dirname + '/listing.html');
        
        //La feuille existe
        else
            res.render(__dirname + '/editor.html', {id_file: req.params.id, nom_file: rows[0].nom,
                rows: rows[0].rows, columns: rows[0].columns});
    });
});


//Bibliothèque Parser
app.get('/parser.js', function (req, res) {
  res.sendFile(__dirname + '/parser.js');
});


//Script javascript pour l'édition
app.get('/editor.js', function (req, res) {
  res.sendFile(__dirname + '/editor.js');
});


//Feuille de style CSS
app.get('/style.css', function (req, res) {
  res.sendFile(__dirname + '/style.css');
});


//Framework jQuery
app.get('/jquery-1.10.2.min.js', function (req, res) {
  res.sendFile(__dirname + '/jquery-1.10.2.min.js');
});


//Lors de l'ouverture d'une connexion
io.on('connection', function (socket) {

    //Réception de l'évènement de l'ajout d'une personne sur une page d'édition
    socket.on('adduser', function (id_file, user) {
        var infos = {id: socket.id, user: user, id_file: id_file};
        users.push(infos);
        updateUsers(id_file);
        io.sockets.emit('updatelist', null);
    });

    //Réception de l'évènement de mise à jour des personnes
    socket.on('update', function (id_file) {
        var users_ = [];
        for (var i in users) {
            if (users[i].id_file == id_file)
                users_.push(users[i]);
        }
        io.to(socket.id).emit('update', users_);
    });

    //Fonction pour récupérer les personnes sur une feuille
    //Elle emet ensuite un signal pour mettre à jour cette liste pour tout le monde
    updateUsers = function(id_file) {
        var users_ = [];
        for (var i in users) {
            if (users[i].id_file == id_file)
                users_.push(users[i]);
        }

        for (var i in users) {
            if (users[i].id_file == id_file)
                io.to(users[i].id).emit('update', users_);
        }
    };

    //Chargement d'une feuille
    socket.on('load', function (id_file) {
        
        //Récupération des styles
        connection.query('SELECT * FROM styles WHERE id_file = ' + parseInt(id_file), function(err, rows, fields) {        
            for (var i in rows) {    
                var style = {cell: rows[i].cell, group: rows[i].group, style: rows[i].style};
                io.to(socket.id).emit('loadstyle', style);
            }
        });

        //Récupération des valeurs
        connection.query('SELECT * FROM cells WHERE id_file = ' + parseInt(id_file), function(err, rows, fields) {         
            for (var i in rows) {
                var cell = {cell: rows[i].cell, input: rows[i].input};
                io.to(socket.id).emit('loadcell', cell);
            }
            io.to(socket.id).emit('load');
        });
    });


    //Mise à jour de la liste des fichiers
    socket.on('list', function () {

        //Récupération des différentes fichiers et du nombre de personnes éditant chacun d'entre eux
        connection.query('SELECT * FROM files ORDER BY id DESC', function(err, rows, fields) {        
            for (var i in rows) { 
                var nb = 0;
                for (var j in users)
                    if (users[j].id_file == rows[i].id)
                        nb++;

                var file = {id: rows[i].id, nom: rows[i].nom, nb: nb};
                io.to(socket.id).emit('list', file);
            }
        });
    });

    //Création d'un nouveau fichier, émission de l'id de cette nouvelle feuille
    socket.on('new', function (nom) {
        var post = {nom: nom};
        connection.query('INSERT INTO files SET ?', post, function(err, result) {
            io.to(socket.id).emit('redirect', result.insertId);
        });
        io.sockets.emit('updatelist', null);
    });

    //Mise à jour des lignes
    socket.on('majrows', function (id_file, rows) {
        connection.query('UPDATE files SET rows = ' + parseInt(rows) + 
            ' WHERE id = ' + parseInt(id_file), null, function(err, result) {
        }); 
    });

    //Mise à jour des colonnes
    socket.on('majcolumns', function (id_file, columns) {
        connection.query('UPDATE files SET columns = ' + parseInt(columns) + 
            ' WHERE id = ' + parseInt(id_file), null, function(err, result) {
        }); 
    });

    //Evènement de déconnexion, on supprimer l'utilisateur de la liste
    socket.on('disconnect', function () {
        var id_file = null;
        for(var i in users) {
            if (users[i].id == socket.id) {
                id_file = users[i].id_file;
                users.splice(i, 1);
            }
        }
        
        updateUsers(id_file);
        io.sockets.emit('updatelist', null);
    });

    //Mise à jour de la cellule pour tout le monde
    socket.on('onblur', function (id_file, cell, val) {
        for (var i in users) {
            if (users[i].id_file == id_file)
                io.to(users[i].id).emit('onblur', socket.id, cell, val);
        }

        connection.query('DELETE FROM cells WHERE `id_file` = ' + parseInt(id_file) + ' AND `cell` = "' + cell + 
            '"', null, function(err, result) {}); 
        if (val != '') {
            var post = {id_file: parseInt(id_file), cell: cell, input: val};
            connection.query('INSERT INTO cells SET ?', post, function(err, result) {}); 
        }
    });

    //Edition en temps réel du contenu de la cellule
    socket.on('onkey', function (id_file, cell, val) {
        for (var i in users) {
            if (users[i].id_file == id_file)
                io.to(users[i].id).emit('onkey', socket.id, cell, val); 
        }
    });

    //Mise à jour du style d'une cellule pour tout le monde
    socket.on('toggle', function (id_file, cell, group, style, dele) {
        for (var i in users) {
            if (users[i].id_file == id_file)
                io.to(users[i].id).emit('toggle', socket.id, cell, group, style); 
        }

        if (group != 'style') {
            connection.query('DELETE FROM styles WHERE `id_file` = ' + parseInt(id_file) + ' AND `cell` = "' + cell + 
                '" AND `group` = "' + group + '"', null, function(err, result) {}); 
        } else if (group == 'style' && dele) {
            connection.query('DELETE FROM styles WHERE `id_file` = ' + parseInt(id_file) + ' AND `cell` = "' + cell + 
                '" AND `group` = "' + group + '" AND `style` = "' + style + '"', null, function(err, result) {}); 
        }

        if (!dele) {
            var post = {id_file: parseInt(id_file), cell: cell, group: group, style: style};
            connection.query('INSERT INTO styles SET ?', post, function(err, result) {}); 
        }
    });

});