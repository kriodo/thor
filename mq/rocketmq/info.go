package rocketmq

import (
	"fmt"
	"sync"
)

// MQManager mq管理
type MQManager struct {
	mu  sync.Map // 唯一标识 - mq信息
	exi sync.Map // 唯一标识 - bool
}

var mqM *MQManager

func init() {
	mqM = &MQManager{}
}

// MQInfo 信息
type MQInfo struct {
	LeaseId   string // 租户标识
	UniqueId  string // 唯一标识
	StartTime int64  // 启动时间
}

func (info *MqConsumerReq) Start() error {
	mutex, _ := mqM.mu.LoadOrStore(info.TopicName, &sync.Mutex{})
	mu := mutex.(*sync.Mutex)
	mu.Lock()
	defer mu.Unlock()
	id := info.FormatUniqueId()
	var exiMap map[string]bool
	if val, exi := mqM.exi.Load(id); exi {
		if exiMap, exi = val.(map[string]bool); exi {
			exiMap[id] = true
		}
	} else {
		exiMap[id] = true
	}
	return nil
}

type IMQ interface {
	FormatUniqueId() string
	Start() error
}

func (info *MqConsumerReq) FormatUniqueId() string {
	return fmt.Sprintf("%s-%s", info.LeaseId, info.TopicName)
}
