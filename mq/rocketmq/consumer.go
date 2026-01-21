package rocketmq

import (
	"context"
	"fmt"
	"github.com/apache/rocketmq-client-go/v2"
	"github.com/apache/rocketmq-client-go/v2/consumer"
	"github.com/apache/rocketmq-client-go/v2/primitive"
)

type MqConsumerReq struct {
	LeaseId      string
	InstanceName string
	TopicName    string
	GroupName    string
	URL          string
	Func         MQSubscribeFunc
}

type MqConsumerResp struct {
	PC rocketmq.PushConsumer
}

type MQSubscribeFunc func(ctx context.Context, msgs ...*primitive.MessageExt) (consumer.ConsumeResult, error)

// StartMQConsumer 启动一个mq
func StartMQConsumer(ctx context.Context, info *MqConsumerReq) error {
	// 增加信息
	//
	client, err := rocketmq.NewPushConsumer(
		consumer.WithInstance(info.InstanceName),
		consumer.WithNameServer([]string{info.URL}),
		consumer.WithGroupName(info.GroupName),
	)
	if err != nil {
		return fmt.Errorf("NewPushConsumer失败：%+v", err)
	}
	err = client.Subscribe(info.TopicName, consumer.MessageSelector{}, info.Func)
	if err != nil {
		return fmt.Errorf("subscribe失败：%+v", err)
	}
	err = client.Start()
	if err != nil {
		return fmt.Errorf("MQ开始失败：%+v", err)
	}
	return nil
}
