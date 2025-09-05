package com.example.transmission.repository;

import com.example.transmission.domain.ProvinceSubscriber;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ProvinceSubscriberRepository extends JpaRepository<ProvinceSubscriber, Long> {
    

}
