class_name Goal
extends Node2D

@onready var back_net_area: Area2D = %BackNetArea

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	back_net_area.body_entered.connect(on_ball_enter_back_net.bind())


func on_ball_enter_back_net(ball:Ball) ->void:
	ball.stop()
